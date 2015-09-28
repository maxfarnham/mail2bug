using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace Mail2Bug.Email.EWS
{
    using Microsoft.AzureAd.Icm.Utility;
    using log4net;

    class EWSMailFolder : IMailFolder
    {
        private readonly Folder _folder;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(EWSMailFolder));

        public EWSMailFolder(Folder folder)
        {
            _folder = folder;
        }

        public int GetTotalCount()
        {
            return _folder.TotalCount;
        }

        public IEnumerable<IIncomingEmailMessage> GetMessages()
        {
            Logger.InfoFormat("Getting email messages...");
            var itemCount = _folder.TotalCount;
            if (itemCount <= 0)
            {
                return new List<IIncomingEmailMessage>();
            }

            var view = new ItemView(itemCount);
            var items = _folder.FindItems(view);
            Logger.InfoFormat("Items found: {0}", items.Count());

            var messages = new List<IIncomingEmailMessage>();
            int junkCount = 0;
            foreach (var item in items)
            {
                if (item is EmailMessage)
                {
                    EWSIncomingMessage message = new EWSIncomingMessage((EmailMessage)item);
                    messages.Add(message);
                }
                else
                {
                    junkCount++;
                    item.Move(WellKnownFolderName.DeletedItems);
                }    
            }

            Logger.InfoFormat("Message count: {0}, Junk count: {1}", messages.Count, junkCount);
            Logger.InfoFormat("Completed getting messages.");
            return messages;
        }
    }
}
