using Microsoft.VisualStudio.TestTools.UnitTesting;

using System.Collections.Generic;

using Mail2Bug.IcmIncidentsApiODataReference;
using Mail2Bug.WorkItemManagement;


namespace Mail2BugUnitTests.WorkItemManagement
{
    [TestClass]
    public class IcmWorkItemManagementUnitTest
    {
        [TestMethod]
        public void IcmIncidentOverridesValidTest()
        {
            // Test to update supported fields with valid values
            var incident = new Incident();
            incident.Keywords = "Initial";
            incident.OwningTeamId = "Initial";
            incident.MonitorId = "Initial";
            incident.ReproSteps = "Initial";
            incident.Severity = 4;
            incident.HitCount = 10;

            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("Keywords", "Changed");
            values.Add("OwningTeamId", "Changed");
            values.Add("MonitorId", "Changed");
            values.Add("ReproSteps", "Changed");
            values.Add("Severity", "3");
            values.Add("HitCount", "5");

            IcmWorkItemManagment.ApplyOverrides(ref incident, values);

            Assert.AreEqual(incident.Keywords, "Changed");
            Assert.AreEqual(incident.OwningTeamId, "Changed");
            Assert.AreEqual(incident.MonitorId, "Changed");
            Assert.AreEqual(incident.ReproSteps, "Changed");
            Assert.AreEqual(incident.Severity, 3);
            Assert.AreEqual(incident.HitCount, 5);
        }

        [TestMethod]
        public void IcmIncidentOverridesInvalidTest()
        {
            // Test attempt to override invalid/unsupported fields & supported fields with invalid values
            var incident = new Incident();
            incident.IsNoise = false;
            RootCauseEntity rce = new RootCauseEntity();
            rce.Title = "Test Root cause";
            incident.RootCause = rce;
            incident.Severity = 4;

            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("RootCause", "Error");
            values.Add("IsNoise", "Error");
            values.Add("Severity", "Error");

            IcmWorkItemManagment.ApplyOverrides(ref incident, values);

            Assert.AreEqual(incident.IsNoise, false);
            Assert.AreEqual(incident.RootCause, rce);
            Assert.AreEqual(incident.Severity, 4);
        }

        [TestMethod]
        public void IcmIncidentOverridesMixedTest()
        {
            // Test mixes valid and invalid overrides - valid overrides should go through despite invalid overrides failing
            var incident = new Incident();
            RootCauseEntity rce = new RootCauseEntity();
            rce.Title = "Test Root cause";
            incident.IsNoise = false;
            incident.RootCause = rce;
            incident.Severity = 4;
            incident.Keywords = "Initial";

            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("RootCause", "Error");
            values.Add("IsNoise", "Error");
            values.Add("Severity", "3");
            values.Add("Keywords", "Changed");

            IcmWorkItemManagment.ApplyOverrides(ref incident, values);

            Assert.AreEqual(incident.IsNoise, false);
            Assert.AreEqual(incident.RootCause, rce);
            Assert.AreEqual(incident.Severity, 3);
            Assert.AreEqual(incident.Keywords, "Changed");
        }
    }
}
