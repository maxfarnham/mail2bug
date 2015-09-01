namespace Mail2BugUnitTests.WorkItemManagement
{
    using System;
    using System.IO;
    using System.Security.Cryptography.X509Certificates;
    using System.Threading.Tasks;

    using Mail2Bug.IcmIncidentsApiODataReference;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using Mail2Bug.WorkItemManagement;

    [TestClass]
    public class IcmODataClientTest
    {
        private  const string ServiceUri = "https://icm.ad.msft.net/api/cert";
        private const string CertThumbprint = "8D565A480BDB7BA78933C009CD13A2B0E5C55CF3";
        private readonly DataServiceODataClient dataServiceClient;
        private const int IncidentId = 11727813;

        public IcmODataClientTest()
        {
            X509Certificate certificate = IcmWorkItemManagment.RetrieveCertificate(CertThumbprint);
            dataServiceClient = new DataServiceODataClient(new Uri(ServiceUri), null, certificate);
        }

        [TestMethod]
        public void GetIncident()
        {
            Incident incident = dataServiceClient.GetIncident(IncidentId);

            Assert.IsNotNull(incident);
            Assert.AreEqual(IncidentId, incident.Id);
            Assert.AreEqual("Email to IcM testing for Authoring team", incident.Title);
        }

        [TestMethod]
        public void ProcessAttachment()
        {
            string dateTimeNow = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string fileName = "FileToAttach" + dateTimeNow + ".txt";
            File.WriteAllText(fileName, "Text file contents " + dateTimeNow);
            Task task = dataServiceClient.ProcessAttachment(IncidentId, fileName);
            task.Wait();
            Assert.IsTrue(task.IsCompleted);
            Assert.AreEqual(TaskStatus.RanToCompletion, task.Status);
        }
    }
}
