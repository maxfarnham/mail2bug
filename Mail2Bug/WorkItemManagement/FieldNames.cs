using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail2Bug.WorkItemManagement
    {
    public static class FieldNames
        {
        public static class ResolutionData
            {
            /// <summary> resolution date field name </summary>
            public const string ResolvedDate = "ResolvedDate";

            /// <summary> contact resolving the incident field name </summary>
            public const string ChangedBy = "ResolvedBy";
            }

        public static class MitigationData
            {
            /// <summary> mitigation date field name </summary>
            public const string MitigatedDate = "MitigatedDate";

            /// <summary> contact mitigating the incident field name </summary>
            public const string MitigatedBy = "MitigatedBy";

            /// <summary> mitigation field name </summary>
            public const string Mitigation = "Mitigation";
            }

        public static class Acknowledgement
            {
            /// <summary> incident acknowledge status field name </summary>
            public const string IsAcknowledged = "IsAcknowledged";

            /// <summary> incident acknowledge date field name </summary>
            public const string AcknowledgeDate = "AcknowledgeDate";

            /// <summary> incident acknowledge contact alias field name </summary>
            public const string AcknowledgeContactAlias = "AcknowledgeContactAlias";
            }
        public static class OccurringLocation
            {
            /// <summary> incident acknowledge status field name </summary>
                public const string DataCenter = "DataCenter";

                /// <summary> incident acknowledge date field name </summary>
                public const string AcknowledgeDate = "AcknowledgeDate";

                /// <summary> incident acknowledge contact alias field name </summary>
                public const string AcknowledgeContactAlias = "AcknowledgeContactAlias";
                
            }

        /// <summary> Incident class </summary>
        public static class Incident
            {
            /// <summary> Incident Id field name </summary>
            public const string Id = "Id";

            /// <summary> Incident create date field name </summary>
            public const string CreateDate = "CreateDate";

            /// <summary> Incident create date field name </summary>
            public const string CreatedBy = "CreateBy";

            /// <summary> Incident modified date field name </summary>
            public const string ModifiedDate = "ModifiedDate";

            /// <summary> Incident severity field name </summary>
            public const string Severity = "Severity";

            /// <summary>Incident Status field name </summary>
            public const string Status = "Status";

            /// <summary> Incident alert source field name </summary>
            public const string Source = "Source";

            /// <summary> Incident correlation id field name </summary>
            public const string CorrelationId = "CorrelationId";

            /// <summary> Incident routing id field name </summary>
            public const string RoutingId = "RoutingId";

            /// <summary> Incident hit count field name </summary>
            public const string HitCount = "HitCount";

            /// <summary> Incident raising location field name </summary>
            public const string RaisingLocation = "RaisingLocation";

            /// <summary> Incident source location field name </summary>
            public const string IncidentLocation = "IncidentLocation";

            /// <summary> owning tenant id field name </summary>
            public const string OwningTenantId = "OwningTenantId";

            /// <summary> owning team id field name </summary>
            public const string OwningTeamId = "OwningTeamId";

            /// <summary> owning contact alias field name </summary>
            public const string OwningContactAlias = "OwningContactAlias";

            /// <summary> incident mitigation data field name </summary>
            public const string MitigationData = "MitigationData";

            /// <summary> incident resolution data field name </summary>
            public const string ResolutionData = "ResolutionData";

            /// <summary> incident customer impacting field name </summary>
            public const string IsCustomerImpacting = "IsCustomerImpacting";

            /// <summary> incident noise field name </summary>
            public const string IsNoise = "IsNoise";

            /// <summary> incident security risk field name </summary>
            public const string IsSecurityRisk = "IsSecurityRisk";

            /// <summary> incident title field name </summary>
            public const string Title = "Title";

            /// <summary> incident repro steps field name </summary>
            public const string ReproSteps = "ReproSteps";

            /// <summary> incident tsgid field name </summary>
            public const string TsgId = "TsgId";

            /// <summary> incident component field name </summary>
            public const string Component = "Component";

            /// <summary> incident customer name field name </summary>
            public const string CustomerName = "CustomerName";

            /// <summary> incident commit date field name </summary>
            public const string CommitDate = "CommitDate";

            /// <summary> incident keywords field name </summary>
            public const string Keywords = "Keywords";

            /// <summary> incident description field name </summary>
            public const string Description = "Description";

            /// <summary> incident description entries field name </summary>
            public const string DescriptionEntries = "DescriptionEntries";

            /// <summary> incident description entries to add field name </summary>
            public const string DescriptionEntriesToAdd = "DescriptionEntriesToAdd";

            /// <summary> acknowledgement data field name </summary>
            public const string AcknowledgementData = "AcknowledgementData";

            /// <summary> impact start date field name </summary>
            public const string ImpactStartDate = "ImpactStartDate";

            /// <summary> incident type field name </summary>
            public const string IncidentType = "IncidentType";

            /// <summary> originating tenant id field name </summary>
            public const string OriginatingTenantId = "OriginatingTenantId";

            /// <summary> parent incident id field name </summary>
            public const string ParentIncidentId = "ParentIncidentId";

            /// <summary> related links count field name </summary>
            public const string RelatedLinksCount = "RelatedLinksCount";

            /// <summary> external links count field name </summary>
            public const string ExternalLinksCount = "ExternalLinksCount";

            /// <summary> last correlation date field name </summary>
            public const string LastCorrelationDate = "LastCorrelationDate";

            /// <summary> child count field name </summary>
            public const string ChildCount = "ChildCount";

            /// <summary> post mortem report data field name </summary>
            public const string PostMortemReportData = "PostMortemReportData";

            /// <summary> reactivation data field name </summary>
            public const string ReactivationData = "ReactivationData";

            /// <summary> subscription id field name </summary>
            public const string SubscriptionId = "SubscriptionId";

            /// <summary> support ticket id field name </summary>
            public const string SupportTicketId = "SupportTicketId";

            /// <summary> monitor id field name </summary>
            public const string MonitorId = "MonitorId";

            /// <summary> incident sub type field name </summary>
            public const string IncidentSubType = "IncidentSubType";

            /// <summary> how fixed field name </summary>
            public const string HowFixed = "HowFixed";

            /// <summary> tsg output field name </summary>
            public const string TsgOutput = "TsgOutput";

            /// <summary> source origin field name </summary>
            public const string Origin = "Origin";


            /// <summary> source origin field name </summary>
            public const string SourceOrigin = "SourceOrigin";

            /// <summary> responsible tenant id field name </summary>
            public const string ResponsibleTenantId = "ResponsibleTenantId";

            /// <summary> impacted services ids field name </summary>
            public const string ImpactedServicesIds = "ImpactedServicesIds";

            /// <summary> impacted teams public ids field name </summary>
            public const string ImpactedTeamsPublicIds = "ImpactedTeamsPublicIds";

            }
        }
    }
