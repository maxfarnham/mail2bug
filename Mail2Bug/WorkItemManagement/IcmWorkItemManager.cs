namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Security.Cryptography.X509Certificates;
    using System.ServiceModel;
    using System.ServiceModel.Security;

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
       
        private ConnectorIncidentManagerClient connectorClient;

        public SortedList<string, int> WorkItemsCache { get; private set; }
        
        public IcmWorkItemManagment(Config.InstanceConfig instanceConfig)
        {
            ValidateConfig(instanceConfig);
            config = instanceConfig;

            X509Certificate certificate = RetrieveCertificate(CertThumbprint);
            dataServiceClient = new DataServiceODataClient(
                new Uri(config.IcmServerConfig.OdataServiceBaseUri),
                config,
                certificate);

            connectorClient = ConnectToIcmInstance();
            Logger.InfoFormat("Connected to IcM");
            if (connectorClient == null)
            {
                Logger.ErrorFormat("Cannot initialize IcM Webservice");
                throw new Exception("Cannot initialize IcM Webservice");
            }

            Logger.InfoFormat("Initializing WorkItems Cache");
            InitTicketCache();
            nameResolver = InitNameResolver();
            dateHolder = DateTime.UtcNow;
        }

        public X509Certificate RetrieveCertificate(string certThumbprint)
        {
            X509Certificate targetCertificate;

            // Get the store where your certificate is in.
            var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);

            store.Open(OpenFlags.ReadOnly);

            // Select your certificate from the store (any way you like).
            X509Certificate2Collection certColl = store.Certificates.Find(
                X509FindType.FindByThumbprint,
                certThumbprint,
                false);

            // Set the certificate property on the container.
            targetCertificate = certColl[0];

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
            var icmConfig = config.IcmServerConfig;

            ValidateConfigString(icmConfig.IcmUri, "IcmServerConfig.IcmUri");
            ValidateConfigString(icmConfig.IcmTenant, "IcmServerConfig.IcmTenant");
            ValidateConfigString(icmConfig.IcmTicketTemplate, "IcmServerConfig.IcmTicketTemplate");
            ValidateConfigString(icmConfig.NamesListFieldName, "IcmServerConfig.NamesListFieldName");
            ValidateConfigString(config.WorkItemSettings.ConversationIndexFieldName, "WorkItemSettings.ConversationIndexFieldName");
        }

        private ConnectorIncidentManagerClient ConnectToIcmInstance()
        {
            try
            {
                Logger.InfoFormat(
                    "Connecting to IcM  {0} using {1} Cert",
                    config.IcmServerConfig.IcmUri, CertThumbprint);
                var icmServer = CreateConnectorClient(config.IcmServerConfig.IcmUri);
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
            Logger.InfoFormat("Initializing Icm Ticket cache");
            WorkItemsCache = new SortedList<string, int>();

            var itemsToCache = dataServiceClient.SearchIncidents(config.IcmServerConfig.TopOption, config.IcmServerConfig.SkipOption, config.IcmServerConfig.FilterOption);
            Logger.InfoFormat("items retrieved by IcM cache query");
            foreach (IcmIncidentsApiODataReference.Incident workItem in itemsToCache)
            {
                try
                {
                    CacheWorkItem((int)workItem.Id);
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

        public void AttachFiles(int workItemId, List<string> fileList)
        {
            foreach (string file in fileList)
            {
                string base64Content = Helpers.FileUtils.FileToString(file);
                dataServiceClient.ProcessAttachment(file, base64Content, workItemId);
                var incident = dataServiceClient.GetIncident(workItemId);
            }
        }

        public void CacheWorkItem(int workItemId)
        {
            var keyField = config.WorkItemSettings.ConversationIndexFieldName;
            if (string.IsNullOrEmpty(keyField))
            {
                Logger.WarnFormat(
                    "Problem caching work item {0}. Field '{1}' is empty - using ID instead.",
                    workItemId,
                    keyField);
                WorkItemsCache[workItemId.ToString(CultureInfo.InvariantCulture)] = workItemId;
            }

            WorkItemsCache[keyField] = workItemId;
        }

        public int CreateWorkItem(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = new AlertSourceIncident();
            incident.ServiceResponsible = new TenantIdentifier("ES Ads Diagnostics");
            if (connectorClient == null)
            {
                connectorClient = ConnectToIcmInstance();
            }
            incident.ImpactStartDate = dateHolder;
            incident.Title = values["Title"];
            incident.Severity = int.Parse(values["Severity"]);
            incident.Description = values["Description"];
            incident.Source = new AlertSourceInfo
            {
                CreateDate = dateHolder,
                CreatedBy = "Mail2IcM",
                IncidentId = "11153837",
                ModifiedDate = dateHolder,
                Origin = values["Origin"]
            };
            incident.OccurringLocation = new IncidentLocation();
                                             //{
                                             //    DataCenter = values["DataCenter"],
                                             //    DeviceGroup = "MyDG",
                                             //    DeviceName = "MyDevice",
                                             //    Environment = values["Environment "],
                                             //    ServiceInstanceId = "AllMine"
                                             //};
            incident.RaisingLocation = new IncidentLocation();
                                           //{
                                           //    DataCenter = "MyDC",
                                           //    DeviceGroup = "MyDG",
                                           //    DeviceName = "MyDevice",
                                           //    Environment = "MyEnv",
                                           //    ServiceInstanceId = "AllMine"
                                           //};
            incident.RoutingId = "Test";
            incident.Status = IncidentStatus.Active;

            const RoutingOptions RoutingOptions = RoutingOptions.None;
            int id = 0;

            try
            {
                // If the following exception is thrown while debugging, open Visual Studio as Administrator.
                // Additional information: Could not establish secure channel for SSL/TLS with authority 'icm.ad.msft.net'.
                IncidentAddUpdateResult result = connectorClient.AddOrUpdateIncident2(config.IcmServerConfig.ConnectorId, incident, RoutingOptions);
                if (result != null)
                {
                    int.TryParse(result.IncidentId.ToString(), out id);
                }
            }
            catch (Exception ex)
            {
                Logger.ErrorFormat("Exception caught while Creating an incident \n{0}", ex);
            }

            return id;
        }

        public void ModifyWorkItem(int workItemId, string comment, Dictionary<string, string> values)
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