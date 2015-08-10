namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Security.Cryptography.X509Certificates;
    using System.ServiceModel;
    using System.ServiceModel.Security;
    using System.Xml.Serialization;

    using log4net;

    using Mail2Bug.ExceptionClasses;
    using Mail2Bug.MessageProcessingStrategies;

    using Microsoft.AzureAd.Icm.Types;
    using Microsoft.AzureAd.Icm.WebService.Client;

    public class IcmWorkItemManagment : IWorkItemManager
    {
        private const string ToolName = "Mail2IcM";
        private const string CertThumbprint = "8D565A480BDB7BA78933C009CD13A2B0E5C55CF3";
        private static readonly ILog Logger = LogManager.GetLogger(typeof(TFSWorkItemManager));
        private readonly Config.InstanceConfig config;
        private readonly DateTime dateHolder;
        private readonly INameResolver nameResolver;
        private readonly DataServiceODataClient dataServiceClient;
        private readonly AlertSourceIncident incidentDefaults;
        private ConnectorIncidentManagerClient connectorClient;

        public SortedList<string, long> WorkItemsCache { get; private set; }

        public IcmWorkItemManagment(Config.InstanceConfig instanceConfig)
        {
            ValidateConfig(instanceConfig);
            config = instanceConfig;
            incidentDefaults = config.IncidentDefaults;

            X509Certificate certificate = RetrieveCertificate(CertThumbprint);
            dataServiceClient = new DataServiceODataClient(
                new Uri(config.IcmClientConfig.OdataServiceBaseUri),
                config,
                certificate);

            connectorClient = ConnectToIcmInstance();
            if (connectorClient == null)
            {
                Logger.ErrorFormat("Cannot initialize IcM Webservice.");
                throw new Exception("Cannot initialize IcM Webservice.");
            }

            Logger.InfoFormat("Connected to IcM.");

            Logger.InfoFormat("Initializing WorkItems Cache.");
            InitTicketCache();
            nameResolver = InitNameResolver();
            dateHolder = DateTime.UtcNow;
        }

        public X509Certificate RetrieveCertificate(string certThumbprint)
        {
            var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadOnly);

            X509Certificate2Collection certColl = store.Certificates.Find(
                X509FindType.FindByThumbprint,
                certThumbprint,
                false);

            X509Certificate targetCertificate = certColl[0];

            store.Close();

            return targetCertificate;
        }

        private static void ValidateConfig(Config.InstanceConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException("config");
            }

            // Temp variable used for shorthand writing below
            var icmConfig = config.IcmClientConfig;

            ValidateConfigString(icmConfig.IcmUri, "IcmClientConfig.IcmUri");
            ValidateConfigString(icmConfig.IcmTenant, "IcmClientConfig.IcmTenant");
            ValidateConfigString(
                config.WorkItemSettings.ConversationIndexFieldName,
                "WorkItemSettings.ConversationIndexFieldName");
        }

        private ConnectorIncidentManagerClient ConnectToIcmInstance()
        {
            try
            {
                Logger.InfoFormat("Connecting to IcM '{0}' using certificate '{1}'.",
                    config.IcmClientConfig.IcmUri, CertThumbprint);
                var icmServer = CreateConnectorClient(config.IcmClientConfig.IcmUri);
                Logger.InfoFormat("Successfully connected to IcM");
                return icmServer;
            }
            catch (Exception ex)
            {
                Logger.WarnFormat("IcM connection attempt failed.\n Exception: {0}", ex);
            }

            return null;
        }

        private static ConnectorIncidentManagerClient CreateConnectorClient(string icmWebServiceBaseUrl)
        {
            WS2007HttpBinding binding = new WS2007HttpBinding(SecurityMode.Transport)
            {
                Name = "IcmBindingConfigCert",
                MaxBufferPoolSize = 4194304,
                MaxReceivedMessageSize = 16777216
            };

            binding.Security.Transport.Realm = string.Empty;
            binding.Security.Transport.ProxyCredentialType = HttpProxyCredentialType.None;
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Certificate;
            binding.Security.Message.EstablishSecurityContext = false;
            binding.Security.Message.NegotiateServiceCredential = true;
            binding.Security.Message.AlgorithmSuite = SecurityAlgorithmSuite.Default;
            binding.Security.Message.ClientCredentialType = MessageCredentialType.Certificate;

            EndpointAddress remoteAddress = new EndpointAddress(icmWebServiceBaseUrl);

            ConnectorIncidentManagerClient client = new ConnectorIncidentManagerClient(binding, remoteAddress);
            if (client.ClientCredentials != null)
            {
                client.ClientCredentials.ClientCertificate.SetCertificate(
                    StoreLocation.LocalMachine,
                    StoreName.My,
                    X509FindType.FindByThumbprint,
                    CertThumbprint);
            }

            return client;
        }

        private void InitTicketCache()
        {
            Logger.InfoFormat("Initializing IcM Ticket cache...");
            WorkItemsCache = new SortedList<string, long>();

            var itemsToCache = dataServiceClient.SearchIncidents(
                config.IcmClientConfig.TopOption,
                config.IcmClientConfig.SkipOption,
                config.IcmClientConfig.FilterOption).ToArray();
            Logger.InfoFormat("{0} items retrieved by IcM cache query.", itemsToCache.Count());
            
            foreach (IcmIncidentsApiODataReference.Incident workItem in itemsToCache)
            {
                try
                {
                    CacheWorkItem(workItem.Id);
                }
                catch (Exception ex)
                {
                    Logger.ErrorFormat("Exception caught while caching Incident with id {0}\n{1}", workItem.Id, ex);
                }
            }
        }

        private static void ValidateConfigString(string value, string configValueName)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new BadConfigException(configValueName);
            }
        }

        public void AttachFiles(long workItemId, List<string> fileList)
        {
            foreach (string file in fileList)
            {
                dataServiceClient.ProcessAttachment(workItemId, file);
            }
        }

        public void CacheWorkItem(long workItemId)
        {
            if (WorkItemsCache.ContainsValue(workItemId))
            {
                return;
            }

            CacheWorkItem(dataServiceClient.GetIncident(workItemId));
        }

        public void CacheWorkItem(IcmIncidentsApiODataReference.Incident workItem)
        {
            var keyField = config.WorkItemSettings.ConversationIndexFieldName;
            if (string.IsNullOrEmpty(keyField))
            {
                Logger.WarnFormat(
                    "Problem caching work item {0}. Field '{1}' is empty - using ID instead.",
                    workItem.Id,
                    keyField);
                WorkItemsCache[workItem.Id.ToString(CultureInfo.InvariantCulture)] = workItem.Id;
            }

            var keyFieldValue = workItem.Keywords;
            Logger.DebugFormat("Work item {0} conversation ID is {1}", workItem.Id, keyFieldValue);
            WorkItemsCache[keyField] = workItem.Id;
        }

        public AlertSourceIncident CreateIncidentWithDefaults(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = new AlertSourceIncident();
            incident.ImpactStartDate = Convert.ToDateTime(values[FieldNames.Incident.CreateDate]);
            incident.Title = values[FieldNames.Incident.Title]; 
            incident.Severity = incidentDefaults.Severity;
            incident.Description = values[FieldNames.Incident.Description];
                //incidentDefaults.Description;
            incident.Keywords = values["ConverstionID"];
            incident.Source = new AlertSourceInfo
            {
                CreatedBy = values[FieldNames.Incident.CreatedBy],
                CreateDate = DateTime.UtcNow,
                IncidentId = Guid.NewGuid().ToString(),
                ModifiedDate = DateTime.UtcNow,
                Origin = incidentDefaults.Source.Origin
            };
            incident.OccurringLocation = new IncidentLocation
            {
                DataCenter = incidentDefaults.OccurringLocation.DataCenter,
                DeviceGroup = incidentDefaults.OccurringLocation.DeviceGroup,
                DeviceName = incidentDefaults.OccurringLocation.DeviceName,
                Environment = incidentDefaults.OccurringLocation.Environment,
                ServiceInstanceId = incidentDefaults.OccurringLocation.ServiceInstanceId
            };
            incident.RaisingLocation = new IncidentLocation
            {
                DataCenter = incidentDefaults.RaisingLocation.DataCenter,
                DeviceGroup = incidentDefaults.RaisingLocation.DeviceGroup,
                DeviceName = incidentDefaults.RaisingLocation.DeviceName,
                Environment = incidentDefaults.RaisingLocation.Environment,
                ServiceInstanceId = incidentDefaults.RaisingLocation.ServiceInstanceId
            };
            incident.RoutingId = config.IcmClientConfig.RoutingId;
            incident.Status = incidentDefaults.Status;

            return incident;
        }

        public long CreateWorkItem(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = this.CreateIncidentWithDefaults(values);
            if (connectorClient == null)
            {
                connectorClient = ConnectToIcmInstance();
            }

            //incident.ServiceResponsible = new TenantIdentifier("ES Ads Diagnostics");

            const RoutingOptions RoutingOptions = RoutingOptions.None;
            long id = 0;

                // If the following exception is thrown while debugging, open Visual Studio as Administrator.
                // Additional information: Could not establish secure channel for SSL/TLS with authority 'icm.ad.msft.net'.
                IncidentAddUpdateResult result = connectorClient.AddOrUpdateIncident2(
                                                                                        config.IcmClientConfig.ConnectorId,
                                                                                        incident,
                                                                                        RoutingOptions);
                if (result != null)
                {
                    id = result.IncidentId.Value;
                    CacheWorkItem(dataServiceClient.GetIncident(id));
                }
            return id;
         }
        public void ModifyWorkItem(long workItemId, string comment, Dictionary<string, string> values)
        {
            if (workItemId <= 0)
            {
                return;
            }

            long incidentId = workItemId;
            var incident = dataServiceClient.GetIncident(incidentId);

            if (incident == null)
            {
                return;
            }

            incident.NewDescriptionEntry = new IcmIncidentsApiODataReference.NewDescriptionEntry
            {
                Text = comment.Replace("\n", "<br>"),
                SubmitDate = dateHolder,
                ChangedBy = values["SenderAlias"]
            };

            incident.Source = new IcmIncidentsApiODataReference.AlertSourceInfo
            {
                IncidentId = incidentId.ToString(),
                Origin = ToolName
            };

            dataServiceClient.UpdateIncident(incident);
        }

        public INameResolver GetNameResolver()
        {
            return nameResolver;
        }

        private IcmNameResolver InitNameResolver()
        {
            var teamlist = new List<string> { "ES Ads Diagnostics" };
            return new IcmNameResolver(teamlist);
        }
    }
}