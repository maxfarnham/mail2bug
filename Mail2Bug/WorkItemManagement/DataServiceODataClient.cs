
using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Linq;
using Mail2Bug.IcmIncidentsApiODataReference;
using Microsoft.Office.Ossm.Search;
using System.Security.Cryptography.X509Certificates;
using Mail2Bug.IcmOnCallApiODataReference.Microsoft.AzureAd.Icm.ODataApi.Models;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using System.IO;
using log4net;

namespace Mail2Bug.WorkItemManagement
    {
    public class DataServiceODataClient : IODataClient
        {
        private Container odataClient;
        private Config.InstanceConfig _config;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(DataServiceODataClient));
        public DataServiceODataClient(Uri serviceEndpointBaseUri, Config.InstanceConfig config)
            {
                this.odataClient = new Container(serviceEndpointBaseUri)
                {
                    IgnoreMissingProperties = true
                };
                _config = config;
            }

        public X509Certificate RetrieveCertificate(string certThumbprint)
            {
            X509Certificate targetCertificate;

            // Get the store where your certificate is in.
            var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);

            store.Open(OpenFlags.ReadOnly);

            // Select your certificate from the store (any way you like).
            X509Certificate2Collection certColl = store.Certificates.Find(X509FindType.FindByThumbprint, certThumbprint, false);

            // Set the certificate property on the container.
            targetCertificate = certColl[0];

            store.Close();

            return targetCertificate;
            }
        public void RegisterEventHandlerForSendingRequest(
            EventHandler<SendingRequest2EventArgs> eventHandler)
            {
            this.odataClient.SendingRequest2 += eventHandler;
            }

        public void UnRegisterEventHandlerForSendingRequest(
            EventHandler<SendingRequest2EventArgs> eventHandler)
            {
            this.odataClient.SendingRequest2 -= eventHandler;
            }

        public Incident GetIncident(long incidentId)
            {
            this.odataClient.MergeOption = MergeOption.OverwriteChanges;
            Incident targetIncident = this.odataClient.incidents.Where(o => o.Id == incidentId).SingleOrDefault();

            if (targetIncident == null)
                {
                throw new Exception("invalid incident provided");
                }

            return targetIncident;
            }
        public void ProcessAttachment(string fileName, string base64Contet,long incidentId)
        {
            IncidentAttachment iattachment = new IncidentAttachment();
            iattachment.Filename = fileName;
            iattachment.ContentsBase64 = base64Contet;
            string json = JsonConvert.SerializeObject(iattachment);

            Logger.DebugFormat("fileName:    {0}", fileName);
            Logger.DebugFormat("base64Contet {0}", base64Contet);

            Logger.DebugFormat("Json String: {0} ", json);

            string createAttachementUri = "https://icm.ad.msft.net/api/user/incidents("+incidentId+")/CreateAttachment";

            Logger.DebugFormat("createAttachementUri {0}", createAttachementUri);

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(createAttachementUri);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            try
                {

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();
                    }
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                    var result = streamReader.ReadToEnd();
                    Logger.DebugFormat("HttpResponse:{0}", result);
                    }
                }
            catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
        }

        public IEnumerable<DescriptionEntry> GetIncidentDescriptionEntries(
            long incidentId,
            string topOption,
            string skipOption)
            {
            string incidentsRouteUri = _config.IcmServerConfig.OdataServiceBaseUri + "/incidents(" + incidentId + ")/DescriptionEntries?$inlinecount=allpages";

            if (string.IsNullOrWhiteSpace(topOption) == false)
                {
                incidentsRouteUri += "&$top=" + topOption;
                }

            if (string.IsNullOrWhiteSpace(skipOption) == false)
                {
                incidentsRouteUri += "&$skip=" + skipOption;
                }

            Uri request = new Uri(incidentsRouteUri);
            var response = (QueryOperationResponse<DescriptionEntry>)odataClient.Execute<DescriptionEntry>(request);

          
         
                foreach (DescriptionEntry description in response)
                    {
                    yield return description;
                    }

                var continuation = response.GetContinuation();
                if (continuation == null)
                    {
                    yield break;
                    }

                response = odataClient.Execute(continuation);

            }
           public void AcknowledgeIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void EditIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void MitigateIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void TransferIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void ResolveIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void ReactivateIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public void UnresolveIncident(Incident incident)
            {
            this.UpdateIncident(incident);
            }

        public IEnumerable<Incident> SearchIncidents(
            string topOption,
            string skipOption,
            string filterQueryOption)
            {
            string incidentsRouteUri = _config.IcmServerConfig.OdataServiceBaseUri + "/incidents?$inlinecount=allpages";

            if (string.IsNullOrWhiteSpace(topOption) == false)
                {
                incidentsRouteUri += "&$top=" + topOption;
                }

            if (string.IsNullOrWhiteSpace(skipOption) == false)
                {
                incidentsRouteUri += "&$skip=" + skipOption;
                }

            if (string.IsNullOrWhiteSpace(filterQueryOption) == false)
                {
                incidentsRouteUri += "&$filter=" + filterQueryOption;
                }
            QueryOperationResponse<Incident> response = null;
            try
                {

                Uri request = new Uri(incidentsRouteUri);
                response = (QueryOperationResponse<Incident>)odataClient.Execute<Incident>(request);
                }
            catch (Exception ex)
                {
                Console.WriteLine(ex.InnerException.ToString());
                }


          
            while (true)
                {
                foreach (Incident incident in response)
                    {
                    yield return incident;
                    }

                var continuation = response.GetContinuation();
                if (continuation == null)
                    {
                    yield break;
                    }

                response = odataClient.Execute(continuation);
                }
            }

        private void UpdateIncident(Incident incident)
            {
            if (incident != null)
                {
                this.odataClient.UpdateObject(incident);
                this.odataClient.SaveChanges(SaveChangesOptions.PatchOnUpdate);
                }
            }
          
        }
       
    }
