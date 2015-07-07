// https://microsoft.sharepoint.com/teams/WAG/EngSys/IncidentManagement/IcM%20User%20Guide/Injecting%20incidents%20programmatically%20into%20IcM%20using%20the%20connector%20model.aspx


namespace Mail2Bug.WorkItemManagement
{
    using System;

    using Microsoft.AzureAd.Icm.Types;
    using Microsoft.AzureAd.Icm.WebService.Client;

    public class IcmDataClient
    {
        // DO NOT CHANGE
        private readonly Guid connectorId = new Guid("416bc155-979d-40d4-b422-e7dbccee36f2");
        
        private ConnectorIncidentManagerClient connectorClient = new ConnectorIncidentManagerClient("prod");

        public void Go()
        {
            AlertSourceIncident incident = new AlertSourceIncident();
            incident.ServiceResponsible = new TenantIdentifier("ES Ads Diagnostics");
            incident.Title = "Try Out ICM for Mail2IcM " + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            incident.Severity = 4;
            incident.Source = new AlertSourceInfo
            {
                CreateDate = DateTime.Now,
                CreatedBy = "brstring",
                IncidentId = Guid.NewGuid().ToString(),
                ModifiedDate = DateTime.Now,
                Origin = "Mail2IcM",
            };
            incident.OccurringLocation = new IncidentLocation
            {
                DataCenter = "MyDC",
                DeviceGroup = "MyDG",
                DeviceName = "MyDevice",
                Environment = "MyEnv",
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
            
            // If the following exception is thrown while debugging, open Visual Studio as Administrator.
            // Additional information: Could not establish secure channel for SSL/TLS with authority 'icm.ad.msft.net'.
            IncidentAddUpdateResult result = connectorClient.AddOrUpdateIncident2(this.connectorId, incident, routingOptions);
        }
    }
}
