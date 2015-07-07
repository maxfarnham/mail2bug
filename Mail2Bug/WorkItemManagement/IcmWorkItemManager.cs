

namespace Mail2Bug.WorkItemManagement
{
    using System;
    using System.Collections.Generic;

    using Mail2Bug.MessageProcessingStrategies;

    public class IcmWorkItemManagment : IWorkItemManager
    {
        private string icmEndpoint = "https://icm.ad.msft.net/api/cert";

        private string certThumbprint = "8D565A480BDB7BA78933C009CD13A2B0E5C55CF3";

        private IcmDataClient icmDataServiceClient = new IcmDataClient();

        public void AttachFiles(int workItemId, List<string> fileList)
        {
            throw new System.NotImplementedException();
        }

        public SortedList<string, int> WorkItemsCache { get; private set; }

        public void CacheWorkItem(int workItemId)
        {
            throw new System.NotImplementedException();
        }

        public int CreateWorkItem(Dictionary<string, string> values)
        {
            throw new System.NotImplementedException();
        }

        public void ModifyWorkItem(int workItemId, string comment, Dictionary<string, string> values)
        {
            throw new System.NotImplementedException();
        }

        public INameResolver GetNameResolver()
        {
            throw new System.NotImplementedException();
        }
    }
}