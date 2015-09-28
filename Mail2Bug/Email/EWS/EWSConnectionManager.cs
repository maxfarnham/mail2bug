using System;
using System.Collections.Generic;
using log4net;
using Microsoft.Exchange.WebServices.Data;

namespace Mail2Bug.Email.EWS
{
    using System.Linq;

    using Microsoft.AzureAd.Icm.Utility;

    /// <summary>
    /// This class caches EWS connection objects based on their settings.
    /// When a caller asks for a new EWS connection, if an appropriate object already exists, we just return that object
    /// thus avoiding the long initialization time for EWS (~1 minute).
    /// 
    /// This works well for configurations where many InstanceConfigs are relying on the same user with different mail folders.
    /// 
    /// Since we may want to be able to turn off caching in some cases, the caching itself is controlled at initialization time.
    /// If caching is disabled, a new connection will be initiated for every call.
    /// </summary>
    public class EWSConnectionManager
    {
        public struct Credentials
        {
            public string EmailAddress;
            public string UserName;
            public string Password;
        }

        public struct EWSConnection
        {
            public ExchangeService Service;
            public RecipientsMailboxManagerRouter Router;
        }

        public EWSConnectionManager(bool enableConnectionCaching)
        {
            _enableConnectionCaching = enableConnectionCaching;

            if (_enableConnectionCaching)
            {
                _cachedConnections = new Dictionary<Tuple<string, string, int>, EWSConnection>();
            }
        }

        public EWSConnection GetConnection(Credentials credentials)
        {
            Logger.InfoFormat("Getting EWS connection for mailbox.");

            EWSConnection ewsConnection;

            if (!_enableConnectionCaching)
            {
                Logger.InfoFormat("Creating non-cached connection for mailbox...");
                ewsConnection = ConnectToEWS(credentials);
            }
            else
            {
                lock (_cachedConnections)
                {
                    var key = GetKeyFromCredentials(credentials);

                    if (_cachedConnections.ContainsKey(key))
                    {
                        Logger.InfoFormat("Connection for '{0}' already exists. Retrieving from cache.", key);
                        ewsConnection = _cachedConnections[key];
                    }
                    else
                    {
                        Logger.InfoFormat("Creating cached connection for '{0}'...", key);
                        _cachedConnections[key] = ConnectToEWS(credentials);
                        ewsConnection = _cachedConnections[key];
                    }
                }
            }

            Logger.InfoFormat("Completed getting connection for mailbox.");
            return ewsConnection;
        }

        private const string ReplyMessage =
            "<body style=\"font-family:sans-serif\"><p>Hello, {0}:</p>" + 
            "<p>The message sent to '{1}' does not match any of the configurations setup to process " + 
            "the message into an IcM ticket. If the email is intended to create an IcM ticket, ensure " + 
            "that the proper configuration is in place on the 'Mail2IcM' server. Otherwise, please " + 
            "remove the address '{1}' from the email recipient list or associated distribution list " +
            "to avoid getting additional notices.</p>" + 
            "<p>Regards,<br/>Mail2IcM Processor</p></body>";

        public void ProcessInboxes()
        {
            foreach (var connection in _cachedConnections)
            {
                Logger.InfoFormat("Processing inbox for connection {0}", connection.Key);
                connection.Value.Router.ProcessInbox();

                ProcessRogueMessages(connection.Value);
            }
        }

        private static void ProcessRogueMessages(EWSConnection ewsConnection)
        {
            Logger.InfoFormat("Processing rogue messages...");
            var folderName = "Rogue Messages";
            var folders = ewsConnection.Service.FindFolders(WellKnownFolderName.Inbox, new FolderView(100));
            var rogueMessageFolder = folders.SingleOrDefault(f => f.DisplayName == folderName);
            if (rogueMessageFolder == null)
            {
                Logger.InfoFormat("Could not find folder '{0}' on server. Using 'Deleted Items' instead.", folderName);
                rogueMessageFolder = Folder.Bind(ewsConnection.Service, WellKnownFolderName.DeletedItems);
            }

            var rogueMessages = ewsConnection.Router.RogueMessages;
            if (rogueMessages.Count == 0)
            {
                Logger.InfoFormat("No rogue messages found.");
                return;
            }

            Logger.InfoFormat("Replying to {0} rogue messages without configuration and moving to '{1}' folder...",
                rogueMessages.Count, rogueMessageFolder.DisplayName);

            string[] previouslyProcessedRogueMessages;
            try
            {
                previouslyProcessedRogueMessages = ewsConnection.Service.FindItems(rogueMessageFolder.Id, new ItemView(1000)).
                    Select(i => new EWSIncomingMessage((EmailMessage)i).ConversationTopic).ToArray();
            }
            catch (Exception e)
            {
                Logger.ErrorFormat("Failed to connect to mailbox server. Will retry on next iteration. {0}", e);
                return;
            }

            foreach (var message in rogueMessages)
            {
                var ewsMessage = message as EWSIncomingMessage;
                if (ewsMessage == null)
                {
                    continue;
                }

                if (previouslyProcessedRogueMessages.Contains(ewsMessage.ConversationTopic))
                {
                    Logger.InfoFormat("Already replied to conversation topic.");
                    MoveMessageToFolder(ewsMessage, rogueMessageFolder);
                    continue;
                }

                var sender = ewsMessage.SenderAlias;
                var recipient = "mail4icm@microsoft.com";

                try
                {
                    Logger.InfoFormat("Replying to '{0}', subject '{1}'.", sender, ewsMessage.Subject);
                    var replyMessage = string.Format(ReplyMessage, sender, recipient);
                    ewsMessage.Reply(replyMessage, true);
                }
                catch (Exception e)
                {
                    Logger.ErrorFormat(
                        "Failed to send reply. Possible cause: original mail item moved by the Exchange Server rules. {0}",
                        e);
                }

                MoveMessageToFolder(ewsMessage, rogueMessageFolder);
            }

            Logger.InfoFormat("Completed replying to rogue messages.");
            rogueMessages.Clear();
        }

        private static void MoveMessageToFolder(EWSIncomingMessage ewsMessage, Folder rogueMessageFolder)
        {
            try
            {
                Logger.InfoFormat("Moving mail item to '{0}' folder.", rogueMessageFolder.DisplayName);
                ewsMessage.MoveMessage(rogueMessageFolder.Id);
            }
            catch (Exception)
            {
                Logger.ErrorFormat(
                    "Failed to move message to '{0}' folder. Possible cause: original mail item moved by the Exchange Server rules.",
                    rogueMessageFolder.DisplayName);
            }
        }

        static private Tuple<string, string, int> GetKeyFromCredentials(Credentials credentials)
        {
            return new Tuple<string, string, int>(
                credentials.EmailAddress,
                credentials.UserName, 
                credentials.Password.GetHashCode());
        }

        static private EWSConnection ConnectToEWS(Credentials credentials)
        {
            Logger.DebugFormat("Initializing FolderMailboxManager for email address {0}", credentials.EmailAddress);
            var exchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
            {
                Credentials = new WebCredentials(credentials.UserName, credentials.Password),
                Timeout = 60000
            };

            exchangeService.AutodiscoverUrl(
                credentials.EmailAddress,
                x =>
                {
                    Logger.DebugFormat("Following redirection for EWS autodiscover: {0}", x);
                    return true;
                }
                );

            Logger.DebugFormat("Service URL: {0}", exchangeService.Url);

            return new EWSConnection()
            {
                Service = exchangeService,
                Router =
                    new RecipientsMailboxManagerRouter(
                        new EWSMailFolder(Folder.Bind(exchangeService,WellKnownFolderName.Inbox)))
            };
        }


        private readonly Dictionary<Tuple<string, string, int>, EWSConnection> _cachedConnections;
        private readonly bool _enableConnectionCaching;

        private static readonly ILog Logger = LogManager.GetLogger(typeof(EWSConnectionManager));
    }
}
