namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Security.Cryptography.X509Certificates;
    using System.ServiceModel;
    using System.ServiceModel.Security;
    using System.Xml;

    using log4net;

    using ExceptionClasses;
    using MessageProcessingStrategies;

    using Microsoft.AzureAd.Icm.Types;
    using Microsoft.AzureAd.Icm.WebService.Client;
    using Microsoft.AzureAd.Icm.XhtmlUtility;

    public class IcmWorkItemManagment : IWorkItemManager
    {
        private const string ToolName = "Mail2IcM";
        private const string CertThumbprint = "8D565A480BDB7BA78933C009CD13A2B0E5C55CF3";

        // IcM description text maximum length is actually 32000 characters but let's keep it 
        // slightly less to avoid boundary problems and leave room for the truncation notice.
        public const int DescriptionLengthMax = 31500;
        public const string TruncationMessage = "*** Description truncated by Mail2IcM ***  See attached email for complete description.";

        private static readonly ILog Logger = LogManager.GetLogger(typeof(IcmWorkItemManagment));
        private readonly Config.InstanceConfig config;
        private readonly DateTime dateHolder;
        private readonly INameResolver nameResolver;
        private readonly DataServiceODataClient dataServiceClient;
        private readonly AlertSourceIncident incidentDefaults;
        private ConnectorIncidentManagerClient connectorClient;

        public SortedList<string, long> WorkItemsCache { get; private set; }

        public IcmWorkItemManagment(Config.InstanceConfig instanceConfig)
        {
            Logger.InfoFormat("Creating IcM work item manager...");
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

            InitWorkItemsCache();

            nameResolver = InitNameResolver();
            dateHolder = DateTime.UtcNow;
            Logger.InfoFormat("Completed creating IcM work item manager.");
        }

        private void InitWorkItemsCache()
        {
            Logger.InfoFormat("Initializing work items cache");

            WorkItemsCache = new SortedList<string, long>();

            var result = dataServiceClient.SearchIncidents(null, null, null);

            foreach (var incident in result)
            {
                if (!string.IsNullOrEmpty(incident.Keywords))
                {
                    if (WorkItemsCache.ContainsKey(incident.Keywords))
                    {
                        Logger.InfoFormat($"Skipping duplicate cache key: {incident.Keywords}");
                        continue;
                    }

                    WorkItemsCache.Add(incident.Keywords, incident.Id);
                }
            }
        }

        public static X509Certificate RetrieveCertificate(string certThumbprint)
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
                Logger.InfoFormat("Connecting to IcM '{0}' using certificate '{1}'...",
                    config.IcmClientConfig.IcmUri, CertThumbprint);
                var icmServer = CreateConnectorClient(config.IcmClientConfig.IcmUri);
                Logger.InfoFormat("Successfully connected to IcM.");
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
            // No operation needed. 
            // IcM incidents will not be cached but rather retrieved live from the IcM service.
            // Retrieving live is a best effort to ensure incident is the most current version of incident. 
        }

        public AlertSourceIncident CreateIncidentWithDefaults(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = new AlertSourceIncident();
            incident.ImpactStartDate = Convert.ToDateTime(values[FieldNames.Incident.CreateDate]);
            incident.Title = values[FieldNames.Incident.Title];
            incident.Severity = incidentDefaults.Severity;

            int maxHyperlinkLength = config.WorkItemSettings.RemoveHyperlinkExceedingNCharacters ?? -1;
            DescriptionEntry descriptionEntry = GenerateDescriptionEntry(values, maxHyperlinkLength);
            incident.DescriptionEntries = new DescriptionEntry[] { descriptionEntry };

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
            if (values["CorrelationId"] != "0")
            {
                incident.CorrelationId = values["CorrelationId"];
            }
            return incident;
        }

        public long CreateWorkItem(Dictionary<string, string> values)
        {
            AlertSourceIncident incident = this.CreateIncidentWithDefaults(values);
            if (connectorClient == null)
            {
                connectorClient = ConnectToIcmInstance();
            }

            long incidentId = 0;
            IncidentAddUpdateResult result = connectorClient.AddOrUpdateIncident2(
                                                                config.IcmClientConfig.ConnectorId,
                                                                incident,
                                                                RoutingOptions.None);

            if (result != null && result.IncidentId.HasValue)
            {
                incidentId = result.IncidentId.Value;
                WorkItemsCache.Add(values["ConverstionID"], incidentId);
            }
            return incidentId;
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

        public static DescriptionEntry GenerateDescriptionEntry(Dictionary<string, string> values, int maxHyperlinkLength)
        {
            DateTime now = DateTime.UtcNow;
            DescriptionTextRenderType renderType = DescriptionTextRenderType.Plaintext;
            string text = values[FieldNames.Incident.Description];

            if (text.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Try to convert from HTML to XHTML, if failed, just leave it as plain text
                string xhtmlValid;
                string errors;
                if (XmlSanitizer.TryMakeXHtml(text, out xhtmlValid, out errors))
                {
                    string xhtmlSanitized;
                    if (XmlSanitizer.SanitizeXml(xhtmlValid, false, false, out xhtmlSanitized, out errors))
                    {
                        text = xhtmlSanitized;
                        renderType = DescriptionTextRenderType.Html;
                    }
                }

                if (!string.IsNullOrEmpty(errors))
                {
                    Logger.Info("Failed to convert message body to HTML. Defaulting to plain text. Conversion message: " + errors);
                }
                else if (text.Length > DescriptionLengthMax)
                {
                    // Truncate string if too long. IcM limits the number of characters in the DescriptionEntry.Text property.
                    text = TruncateXml(text, DescriptionLengthMax, maxHyperlinkLength);
                }
            }
            else if (text.Length > DescriptionLengthMax)
            {
                // Truncate string if too long. IcM limits the number of characters in the DescriptionEntry.Text property.
                text = text.Substring(0, DescriptionLengthMax);
                text += TruncationMessage;
            }


            var descriptionEntry = new DescriptionEntry
            {
                Cause = DescriptionEntryCause.Created,
                Date = now,
                ChangedBy = values[FieldNames.Incident.CreatedBy],
                SubmitDate = now,
                SubmittedBy = values[FieldNames.Incident.CreatedBy],
                Text = text,
                RenderType = renderType,
            };

            return descriptionEntry;
        }

        public static string TruncateXml(string xml, int maxLength, int maxHyperlinkLength)
        {
            XmlDocument document = new XmlDocument();
            document.LoadXml(xml);

            Stack<XmlNode> nodeStack = new Stack<XmlNode>();
            XmlNode currentNode = document.FirstChild;
            int accumulatedLength = 0;
            while (currentNode != null)
            {
                if ((currentNode.Name == "a") && (maxHyperlinkLength >= 0) && (currentNode.Attributes != null))
                {
                    XmlNode hrefNode = currentNode.Attributes.GetNamedItem("href");
                    if (hrefNode?.Value != null && (hrefNode.Value.Length > maxHyperlinkLength) && (currentNode.ParentNode != null))
                    {
                        XmlNode newNode = document.CreateNode(XmlNodeType.Element, "div", "");
                        newNode.InnerText = "** Mail2IcM removed hyperlink **";
                        currentNode.ParentNode.ReplaceChild(newNode, currentNode);
                        currentNode = newNode;
                    }
                }

                int lengthIncrease = currentNode.OuterXml.Length - currentNode.InnerXml.Length;
                if ((accumulatedLength + lengthIncrease) > maxLength)
                {
                    XmlNode nodeToDelete = currentNode;

                    while (nodeStack.Count > 0)
                    {
                        currentNode = nodeStack.Pop();

                        while (nodeToDelete != null)
                        {
                            XmlNode nextNode = nodeToDelete.NextSibling;
                            currentNode.RemoveChild(nodeToDelete);
                            nodeToDelete = nextNode;
                        }

                        nodeToDelete = currentNode.NextSibling;
                    }

                    currentNode = currentNode.NextSibling;
                    continue;
                }

                accumulatedLength += lengthIncrease;

                XmlNode child = currentNode.FirstChild;
                if (child != null)
                {
                    nodeStack.Push(currentNode);
                    currentNode = child;
                    continue;
                }

                XmlNode sibling = currentNode.NextSibling;
                if (sibling != null)
                {
                    currentNode = sibling;
                    continue;
                }

                while ((nodeStack.Count > 0) && (nodeStack.Peek().LastChild == currentNode))
                {
                    currentNode = nodeStack.Pop();
                }

                currentNode = currentNode.NextSibling;
            }

            if (xml.Length != document.OuterXml.Length)
            {
                XmlNode truncateMessageNode = document.CreateNode(XmlNodeType.Element,  "div", "");
                truncateMessageNode.InnerText = TruncationMessage;
                document.LastChild.AppendChild(truncateMessageNode);
            }

            string result;
            XmlWriterSettings settings = new XmlWriterSettings { OmitXmlDeclaration = true };
            using (var stringWriter = new StringWriter())
            using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
            {
                document.WriteTo(xmlWriter);
                xmlWriter.Flush();
                result = stringWriter.ToString();
            }

            return result;
        }
    }
}
