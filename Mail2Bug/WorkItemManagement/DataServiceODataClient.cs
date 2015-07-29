namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;
    using System.Data.Services.Client;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;

    using Mail2Bug.IcmIncidentsApiODataReference;
    using Mail2Bug.IcmOnCallApiODataReference.Microsoft.AzureAd.Icm.ODataApi.Models;

    using Microsoft.AzureAd.Icm.Types;

    using Newtonsoft.Json;

    using DescriptionEntry = Mail2Bug.IcmIncidentsApiODataReference.DescriptionEntry;
    using Incident = Mail2Bug.IcmIncidentsApiODataReference.Incident;

    public class DataServiceODataClient
    {
        private readonly Container odataClient;
        private readonly X509Certificate certificate;
        private readonly Config.InstanceConfig config;

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
            Incident targetIncident = this.odataClient.incidents.SingleOrDefault(o => o.Id == incidentId);

            if (targetIncident == null)
            {
                throw new Exception("invalid incident provided");
            }

            return targetIncident;
        }

        public void ProcessAttachment(string fileName, string base64Contet, long incidentId)
        {
            Attachment attachment = new Attachment(fileName, new MemoryStream(Encoding.UTF8.GetBytes(base64Contet)));

            IncidentAttachment iattachment = new IncidentAttachment();
            iattachment.Filename = fileName;
            iattachment.ContentsBase64 = base64Contet;
            string json = JsonConvert.SerializeObject(attachment);
            string createAttachementUri = "https://icm.ad.msft.net/api/cert/incidents(" + incidentId
                                          + ")/CreateAttachment";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(createAttachementUri);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";
            //httpWebRequest.ContentLength = json.Length;

            try
            {
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    streamWriter.Write(json);
                    streamWriter.Flush();
                }

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
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
            string incidentsRouteUri = this.config.IcmServerConfig.OdataServiceBaseUri + "/incidents(" + incidentId
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
            string incidentsRouteUri = this.config.IcmServerConfig.OdataServiceBaseUri + "/incidents?$inlinecount=allpages";

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
