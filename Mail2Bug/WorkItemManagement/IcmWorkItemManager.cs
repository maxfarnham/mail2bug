

namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using log4net;
    using Mail2Bug.MessageProcessingStrategies;
    using Mail2Bug.ExceptionClasses;
    using Microsoft.AzureAd.Icm.WebService.Client;
    using System.ServiceModel;
    using System.Security.Cryptography.X509Certificates;
    using System.ServiceModel.Security;
    using Microsoft.AzureAd.Icm.Types;
    using System.Data.Services.Client;
    using System.Globalization;
    using IcmOnCallApiODataReference.Microsoft.AzureAd.Icm.ODataApi.Models;
    

    public class IcmWorkItemManagment : IWorkItemManager
    {
        private const string _toolName = "Mail2ICM";
        private readonly Config.InstanceConfig _config;
        private readonly DateTime _dateHolder;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(TFSWorkItemManager));
        private readonly INameResolver _nameResolver;
        private static string _certThumbprint = "8D565A480BDB7BA78933C009CD13A2B0E5C55CF3";
        private ConnectorIncidentManagerClient _connectorClient;
        public SortedList<string, int> WorkItemsCache { get; private set; }
        private  DataServiceODataClient dataServiceClient ;
        private   CertAuthODataClient certAuthODataClient ;
        private X509Certificate certificate;
           // X509Certificate certificate = CommonUtils.RetrieveCertificate(AppConfig.CertThumbprint);
            //certAuthODataClient.ClientCertificate = certificate;


        public IcmWorkItemManagment (Config.InstanceConfig config)
        {
            
            ValidateConfig(config);
            _config = config;
            
            dataServiceClient= new DataServiceODataClient(new Uri (_config.IcmServerConfig.OdataServiceBaseUri),_config);
            certAuthODataClient = new CertAuthODataClient(dataServiceClient);
            certificate = dataServiceClient.RetrieveCertificate(_certThumbprint);
            certAuthODataClient.ClientCertificate=certificate;

            _connectorClient = ConnectToICMInstance();
            Logger.InfoFormat("Connected to ICM");
            if (_connectorClient == null)
            {
                Logger.ErrorFormat("Cannot initialize ICM Webservice");
                throw new Exception("Cannot initialize ICM Webservice");
            }
            Logger.InfoFormat("Geting TFS Project");
            Logger.InfoFormat("Initializing WorkItems Cache");
            InitTicketCache();
            _nameResolver = InitNameResolver();
            _dateHolder = DateTime.UtcNow;
        }

        private static void ValidateConfig(Config.InstanceConfig config)
        {
            if (config == null) throw new ArgumentNullException("config");

            // Temp variable used for shorthand writing below
            var icmConfig = config.IcmServerConfig;

            ValidateConfigString(icmConfig.IcmUri, "IcmServerConfig.IcmUri");
            ValidateConfigString(icmConfig.Certificate, "IcmServerConfig.Certificate");
            ValidateConfigString(icmConfig.IcmTenant, "IcmServerConfig.IcmTenant");
            ValidateConfigString(icmConfig.IcmTicketTemplate, "IcmServerConfig.IcmTicketTemplate");
            ValidateConfigString(icmConfig.NamesListFieldName, "IcmServerConfig.NamesListFieldName");
            ValidateConfigString(config.WorkItemSettings.ConversationIndexFieldName, "WorkItemSettings.ConversationIndexFieldName");
        }
        private ConnectorIncidentManagerClient ConnectToICMInstance()
            {
            try
                {
                Logger.InfoFormat("Connecting to ICM  {0} using {1} Cert", _config.IcmServerConfig.IcmUri, _config.IcmServerConfig.Certificate);
                var icmServer = CreateConnectorClient(_config.IcmServerConfig.IcmUri);
                Logger.InfoFormat("Successfully connected to ICM");
                return icmServer;
                }
            catch (Exception ex)
                {
                Logger.WarnFormat("ICM connection attempt failed.\n Exception: {0}", ex);
                }
            return null;
            }
     private static ConnectorIncidentManagerClient CreateConnectorClient(string icmWebServiceBaseUrl)
       {
        ConnectorIncidentManagerClient client;
        WS2007HttpBinding binding;
        EndpointAddress remoteAddress;
        binding = new WS2007HttpBinding(SecurityMode.Transport)
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
         remoteAddress = new EndpointAddress(icmWebServiceBaseUrl);
         client = new ConnectorIncidentManagerClient(binding, remoteAddress);
         if (client.ClientCredentials != null)
         {
             client.ClientCredentials.ClientCertificate.SetCertificate(StoreLocation.LocalMachine, 
                                                                       StoreName.My, 
                                                                       X509FindType.FindByThumbprint, 
                                                                        _certThumbprint);
         }
         return client;
        }
     private void InitTicketCache()
       {
         Logger.InfoFormat("Initializing Icm Ticket cache");
         WorkItemsCache = new SortedList<string, int>();


         //var itemsToCache = SearchIncidents(certAuthODataClient) as ICollection<Mail2Bug.IcmIncidentsApiODataReference.Incident>;
         //search TFS to get list
         var itemsToCache = SearchIncidents(certAuthODataClient);
         Logger.InfoFormat("items retrieved by ICM cache query");
         foreach (Mail2Bug.IcmIncidentsApiODataReference.Incident workItem in itemsToCache)
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
            if (string.IsNullOrEmpty(value)) throw new BadConfigException(configValueName);
        }

        public void AttachFiles(int workItemId, List<string> fileList)
        {
            foreach ( string file in fileList)
            {
                string base64Content = Helpers.FileUtils.FileToString(file);
                dataServiceClient.ProcessAttachment(file, base64Content,workItemId);
            }
        }
        public void CacheWorkItem(int workItemId)
        {
            var keyField = _config.WorkItemSettings.ConversationIndexFieldName;
                if (string.IsNullOrEmpty(keyField))
                {
                    Logger.WarnFormat("Problem caching work item {0}. Field '{1}' is empty - using ID instead.", workItemId, keyField);
                    WorkItemsCache[workItemId.ToString(CultureInfo.InvariantCulture)] = workItemId;
                }
            WorkItemsCache[keyField] = workItemId;
        }

        public int CreateWorkItem(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = new AlertSourceIncident();
            incident.ServiceResponsible = new TenantIdentifier("ES Ads Diagnostics");
            if( _connectorClient==null)  _connectorClient= ConnectToICMInstance();

            incident.ImpactStartDate = _dateHolder;
            incident.Title = values["Title"]; 
            incident.Severity = int.Parse( values["Severity"]);
            incident.Description = values["Description"];
            incident.Source = new AlertSourceInfo
            {
                CreateDate = _dateHolder,
                CreatedBy = "tjimma",
                IncidentId = "11153837",
                ModifiedDate = _dateHolder,
                Origin = values["Orgin"]
            };
            incident.OccurringLocation = new IncidentLocation
            {
                DataCenter = values["DataCenter"],
                DeviceGroup = "MyDG",
                DeviceName = "MyDevice",
                Environment =  values["Environment "],
                ServiceInstanceId = "AllMine"
            };
            incident.RaisingLocation = new IncidentLocation
            {
                DataCenter = "MyDC",
                DeviceGroup = "MyDG",
                DeviceName = "MyDevice",
                Environment = "MyEnv",
                ServiceInstanceId = "AllMine"
            };
            incident.RoutingId = "Test";
            incident.Status = IncidentStatus.Active;

            
            
            RoutingOptions routingOptions = RoutingOptions.None;
            int id = 0;
            // If the following exception is thrown while debugging, open Visual Studio as Administrator.
            // Additional information: Could not establish secure channel for SSL/TLS with authority 'icm.ad.msft.net'.

            try
            {
                 IncidentAddUpdateResult result = _connectorClient.AddOrUpdateIncident2(new Guid(_config.IcmServerConfig.Certificate), incident, routingOptions);
                 if ( result!=null)
                 {
                     int.TryParse(result.IncidentId.ToString(),out id);
                 }
            }
            catch (Exception ex)
            {
                Logger.ErrorFormat("Exception caught while Creating an incident \n{0}", ex);
            }
            return id;
        }
        private IEnumerable<Mail2Bug.IcmIncidentsApiODataReference.Incident> SearchIncidents(IODataClient client)
        {
            string filterOption;
            filterOption = _config.IcmServerConfig.FilterOption;
            string topOption = _config.IcmServerConfig.TopOption;
            string skipOption = _config.IcmServerConfig.SkipOption;
            return client.SearchIncidents(topOption, skipOption, filterOption);
        }
        public void ModifyWorkItem(int workItemId, string comment, Dictionary<string, string> values)
        {
            if (workItemId <= 0) return;
            long incidentId = (long) workItemId;
            var incident = certAuthODataClient.GetIncident(incidentId);

            
            if (incident == null) return;
            incident.NewDescriptionEntry = new IcmIncidentsApiODataReference.NewDescriptionEntry
             {
                 Text = comment.Replace("\n", "<br>"),
                 SubmitDate = _dateHolder,
                 ChangedBy = values["SenderAlias"]
             };

            incident.Source = new IcmIncidentsApiODataReference.AlertSourceInfo
            {
                IncidentId = incidentId.ToString(),
                Origin = _toolName
            };
            certAuthODataClient.EditIncident(incident);
        }

        public INameResolver GetNameResolver()
        {
            return _nameResolver;
        }
        private IcmNameResolver InitNameResolver()
        {
           var teamlist = new List<string>();
           teamlist.Add("ES Ads Diagnostics");
           return new IcmNameResolver(teamlist);
        }
    }
}