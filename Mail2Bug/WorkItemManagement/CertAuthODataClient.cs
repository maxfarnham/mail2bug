using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using Mail2Bug.IcmIncidentsApiODataReference;
using System.Data.Services.Client;

namespace Mail2Bug.WorkItemManagement
    {
    public class CertAuthODataClient :IODataClient
        {
        private X509Certificate clientCertificate = null;

        private IODataClient odataClient;

        public CertAuthODataClient( IODataClient client)
            {
                this.odataClient = client;
            }

        public X509Certificate ClientCertificate
            {
            get
                {
                return clientCertificate;
                }
            set
                {
                if (value == null)
                    {
                    // if the event has been hooked up before, we should remove it
                    if (clientCertificate != null)
                        {
                        this.UnRegisterEventHandlerForSendingRequest(this.OnSendingRequest_AddCertificate);
                        }
                    // this.SendingRequest -= this.OnSendingRequest_AddCertificate;
                    }
                else
                    {
                    // hook up the event if its being set to something non-null
                    if (clientCertificate == null)
                        {
                        this.RegisterEventHandlerForSendingRequest(this.OnSendingRequest_AddCertificate);
                        }
                    // this.SendingRequest += this.OnSendingRequest_AddCertificate;
                    }

                clientCertificate = value;
                }
            }

        public void RegisterEventHandlerForSendingRequest(
            EventHandler<SendingRequest2EventArgs> eventHandler)
            {
            odataClient.RegisterEventHandlerForSendingRequest(eventHandler);
            }

        public void UnRegisterEventHandlerForSendingRequest(
            EventHandler<SendingRequest2EventArgs> eventHandler)
            {
            odataClient.UnRegisterEventHandlerForSendingRequest(eventHandler);
            }

        public Incident GetIncident(long incidentId)
            {
            return odataClient.GetIncident(incidentId);
            }


        public void AcknowledgeIncident(Incident incident)
            {
            odataClient.AcknowledgeIncident(incident);
            }

        public void EditIncident(Incident incident)
            {
            odataClient.EditIncident(incident);
            }

        public void MitigateIncident(Incident incident)
            {
            odataClient.MitigateIncident(incident);
            }

        public void TransferIncident(Incident incident)
            {
            odataClient.TransferIncident(incident);
            }

        public void ResolveIncident(Incident incident)
            {
            odataClient.ResolveIncident(incident);
            }

        public void ReactivateIncident(Incident incident)
            {
            odataClient.ReactivateIncident(incident);
            }

        public void UnresolveIncident(Incident incident)
            {
            odataClient.UnresolveIncident(incident);
            }

        public IEnumerable<Incident> SearchIncidents(
            string topOption,
            string skipOption,
            string filterQueryOption)
            {
            return odataClient.SearchIncidents(topOption, skipOption, filterQueryOption);
            }

        public IEnumerable<DescriptionEntry> GetIncidentDescriptionEntries(
            long incidentId,
            string topOption,
            string skipOption)
            {
            return odataClient.GetIncidentDescriptionEntries(incidentId, topOption, skipOption);
            }

        private void OnSendingRequest_AddCertificate(
            object sender,
            SendingRequest2EventArgs args)
            {

            if (null != ClientCertificate)
                {
                ((HttpWebRequestMessage)args.RequestMessage).HttpWebRequest.ClientCertificates.Add(ClientCertificate);
                }
            }
        }
    }
