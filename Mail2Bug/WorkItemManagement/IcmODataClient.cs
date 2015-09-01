namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using System.Data.Services.Client;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;
    using System.Threading.Tasks;

    using log4net;

    using Mail2Bug.IcmIncidentsApiODataReference;
    using Mail2Bug.IcmOnCallApiODataReference.Microsoft.AzureAd.Icm.ODataApi.Models;

    using Microsoft.AzureAd.Icm.Types;
    using Microsoft.Exchange.WebServices.Data;

    using Newtonsoft.Json;

    using Attachment = Microsoft.AzureAd.Icm.Types.Attachment;
    using DescriptionEntry = Mail2Bug.IcmIncidentsApiODataReference.DescriptionEntry;
    using Incident = Mail2Bug.IcmIncidentsApiODataReference.Incident;
    using Task = System.Threading.Tasks.Task;

    public class DataServiceODataClient
    {
        private readonly Container odataClient;
        private readonly X509Certificate certificate;
        private readonly Config.InstanceConfig config;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(DataServiceODataClient));

        public DataServiceODataClient(Uri serviceEndpointBaseUri, Config.InstanceConfig config, X509Certificate certificate)
        {
            this.odataClient = new Container(serviceEndpointBaseUri) { IgnoreMissingProperties = true };
            this.odataClient.SendingRequest2 += OdataClientOnSendingRequest2;

            this.config = config;
            this.certificate = certificate;
        }

        private void OdataClientOnSendingRequest2(object sender, SendingRequest2EventArgs sendingRequest2EventArgs)
        {
            ((HttpWebRequestMessage)sendingRequest2EventArgs.RequestMessage).HttpWebRequest.ClientCertificates.Add(this.certificate);
        }

        public void RegisterEventHandlerForSendingRequest(EventHandler<SendingRequest2EventArgs> eventHandler)
        {
            this.odataClient.SendingRequest2 += eventHandler;
        }

        public void UnRegisterEventHandlerForSendingRequest(EventHandler<SendingRequest2EventArgs> eventHandler)
        {
            this.odataClient.SendingRequest2 -= eventHandler;
        }

        public Incident GetIncident(long incidentId)
        {
            this.odataClient.MergeOption = MergeOption.OverwriteChanges;
            Incident targetIncident = this.odataClient.incidents.Where(o => o.Id == incidentId).SingleOrDefault();

            if (targetIncident == null)
            {
                throw new Exception("Invalid incident provided.");
            }

            return targetIncident;
        }

        public class AttachmentWrapper
        {
            public IncidentAttachment Attachment { get; set; }
        }
 
        public async Task ProcessAttachment(long incidentId, string fileName)
        {
            String fileAsBase64 = Convert.ToBase64String(File.ReadAllBytes(fileName));
            AttachmentWrapper payload = new AttachmentWrapper();
            payload.Attachment = IncidentAttachment.CreateIncidentAttachment(Path.GetFileName(fileName), fileAsBase64);

            string createAttachementUri = "https://icm.ad.msft.net/api/cert/incidents(" + incidentId
                                          + ")/CreateAttachment";
            Logger.DebugFormat("CreateAttachementUri {0}", createAttachementUri);

            try
            {
                using (WebRequestHandler handler = new WebRequestHandler())
                {
                    handler.ClientCertificates.Add(certificate);

                    using (HttpClient client = new HttpClient(handler))
                    {
                        client.BaseAddress = new Uri(createAttachementUri);
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                        var response = await client.PostAsJsonAsync(createAttachementUri, payload);
                        Logger.InfoFormat(response.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.ErrorFormat(ex.ToString());
            }
        }

        public IEnumerable<DescriptionEntry> GetIncidentDescriptionEntries(
            long incidentId,
            string topOption,
            string skipOption)
        {
            string incidentsRouteUri = this.config.IcmClientConfig.OdataServiceBaseUri + "/incidents(" + incidentId
                                       + ")/DescriptionEntries?$inlinecount=allpages";

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

            odataClient.Execute(continuation);
        }

        public IEnumerable<Incident> SearchIncidents(string topOption, string skipOption, string filterQueryOption)
        {
            string incidentsRouteUri = this.config.IcmClientConfig.OdataServiceBaseUri + "/incidents?$inlinecount=allpages";

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

        public void UpdateIncident(Incident incident)
        {
            if (incident != null)
            {
                this.odataClient.UpdateObject(incident);
                this.odataClient.SaveChanges(SaveChangesOptions.PatchOnUpdate);
            }
        }
    }
}
