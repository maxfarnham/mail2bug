namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Data.Services.Client;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Security.Cryptography.X509Certificates;

    using log4net;

    using IcmIncidentsApiODataReference;
    using IcmOnCallApiODataReference.Microsoft.AzureAd.Icm.ODataApi.Models;
    using Incident = IcmIncidentsApiODataReference.Incident;
    using Task = System.Threading.Tasks.Task;

    public class DataServiceODataClient
    {
        private readonly Container odataClient;
        private readonly X509Certificate certificate;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(DataServiceODataClient));

        public DataServiceODataClient(Uri serviceEndpointBaseUri, X509Certificate certificate)
        {
            odataClient = new Container(serviceEndpointBaseUri) { IgnoreMissingProperties = true };
            odataClient.SendingRequest2 += OdataClientOnSendingRequest2;

            this.certificate = certificate;
        }

        private void OdataClientOnSendingRequest2(object sender, SendingRequest2EventArgs sendingRequest2EventArgs)
        {
            ((HttpWebRequestMessage)sendingRequest2EventArgs.RequestMessage).HttpWebRequest.ClientCertificates.Add(this.certificate);
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
