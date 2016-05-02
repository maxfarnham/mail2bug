using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.AzureAd.Icm.Types;

using System.Collections.Generic;

using Mail2Bug.IcmIncidentsApiODataReference;
using Mail2Bug.WorkItemManagement;


namespace Mail2BugUnitTests.WorkItemManagement
{
    [TestClass]
    public class IcmWorkItemManagementUnitTest
    {
        [TestMethod]
        public void IcmIncidentOverridesValuesTest()
        {
            // Tests mix of valid & invalid values and atrributes
            // Invalid values & fields should be ignored with no side effects, valid fields with valid values should be updated
            var incident = new Mail2Bug.IcmIncidentsApiODataReference.Incident();
            incident.IsNoise = false;
            Mail2Bug.IcmIncidentsApiODataReference.RootCauseEntity rce = new Mail2Bug.IcmIncidentsApiODataReference.RootCauseEntity();
            rce.Title = "Test Root cause";
            incident.RootCause = rce;
            incident.Severity = 4;
            incident.HitCount = 10;
            incident.Keywords = "Initial";
            incident.OwningTeamId = "Initial";

            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("RootCause", "Error");
            values.Add("IsNoise", "Error");
            values.Add("HitCount", "Error");
            values.Add("DoesntExist", "Error");
            values.Add("Keywords", "Changed");
            values.Add("OwningTeamId", "Changed");
            values.Add("Severity", "3");

            IcmWorkItemManagment.ApplyOverrides(ref incident, values);

            Assert.AreEqual(incident.IsNoise, false);
            Assert.AreEqual(incident.RootCause, rce);
            Assert.AreEqual(incident.Keywords, "Changed");
            Assert.AreEqual(incident.OwningTeamId, "Changed");
            Assert.AreEqual(incident.Severity, 3);
            Assert.AreEqual(incident.HitCount, 10);
        }

        [TestMethod]
        public void IcmIncidentOverridesTypeTest()
        {
            // Test how overrides handle different types passed as "incident"
            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("RootCause", "Error");
            values.Add("IsNoise", "Error");
            values.Add("Severity", "3");
            values.Add("Keywords", "Changed");

            // These three should fail without error
            string testString = "";
            IcmWorkItemManagment.ApplyOverrides(ref testString, values);
            int testInt = 0;
            IcmWorkItemManagment.ApplyOverrides(ref testInt, values);
            Mail2Bug.IcmIncidentsApiODataReference.IncidentImpactedComponent testComplexObj = new Mail2Bug.IcmIncidentsApiODataReference.IncidentImpactedComponent();
            IcmWorkItemManagment.ApplyOverrides(ref testComplexObj, values);

            // These two should be handles identically
            Mail2Bug.IcmIncidentsApiODataReference.Incident testIncident = new Mail2Bug.IcmIncidentsApiODataReference.Incident();
            testIncident.Keywords = "Initial";
            testIncident.Severity = 4;
            IcmWorkItemManagment.ApplyOverrides(ref testIncident, values);
            Assert.AreEqual(testIncident.Severity, 3);
            Assert.AreEqual(testIncident.Keywords, "Changed");

            AlertSourceIncident testAlertSourceIncident = new AlertSourceIncident();
            testAlertSourceIncident.Keywords = "Initial";
            testAlertSourceIncident.Severity = 4;
            IcmWorkItemManagment.ApplyOverrides(ref testAlertSourceIncident, values);
            Assert.AreEqual(testAlertSourceIncident.Severity, 3);
            Assert.AreEqual(testAlertSourceIncident.Keywords, "Changed");
        }
    }
}
