using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Services.Client;
using Mail2Bug.IcmIncidentsApiODataReference;

namespace Mail2Bug.WorkItemManagement
    {
    public interface IODataClient
        {
        void RegisterEventHandlerForSendingRequest(EventHandler<SendingRequest2EventArgs> eventHandler);

        void UnRegisterEventHandlerForSendingRequest(EventHandler<SendingRequest2EventArgs> eventHandler);

        Incident GetIncident(long incidentId);

        void AcknowledgeIncident(Incident incident);

        void EditIncident(Incident incident);

        void MitigateIncident(Incident incident);

        void TransferIncident(Incident incident);

        void ResolveIncident(Incident incident);

        void ReactivateIncident(Incident incident);

        void UnresolveIncident(Incident incident);

        IEnumerable<Incident> SearchIncidents(
            string topOption,
            string skipOption,
            string filterQueryOption);

        IEnumerable<DescriptionEntry> GetIncidentDescriptionEntries(
            long incidentId,
            string topOption,
            string skipOption);
        }
    }
