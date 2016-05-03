namespace Mail2Bug.Email.EWS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using log4net;

    using Microsoft.Applications.Telemetry;

    public class RecipientsMailboxManagerRouter
    {
        private readonly ILogger logger = Microsoft.Applications.Telemetry.Server.LogManager.GetLogger();

        public delegate bool MessageEvaluator(IIncomingEmailMessage message);

        public RecipientsMailboxManagerRouter(IMailFolder folder)
        {
            _folder = folder;
        }

        public int RegisterMailbox(MessageEvaluator evaluator)
        {
            var id = _nextId++;
            _clients[id] = new ClientData
            {
                Messages = new List<IIncomingEmailMessage>(),
                Evaluator = evaluator
            };

            return id;
        }

        public IEnumerable<IIncomingEmailMessage> GetMessages(int clientId)
        {
            if (!_clients.ContainsKey(clientId))
            {
                Logger.ErrorFormat("Can't retrieve messages for client ID {0}. No such client registered",
                    clientId);

                return new List<IIncomingEmailMessage>();
            }

            Logger.DebugFormat("Getting messages for client {0}", clientId);
            var incomingEmailMessages = _clients[clientId].Messages;
            Logger.DebugFormat("{0} messages found for client ID {1}", incomingEmailMessages.Count, clientId);

            return incomingEmailMessages;
        }

        public void ProcessInbox()
        {
            Logger.Info("Processing inbox for RecipientsMailboxManagerRouter");

            if (_clients.Count == 0)
            {
                Logger.Info("No clients registered for RecipientsMailboxManagerRouter - returning");
                return;
            }

            foreach (var clientData in _clients)
            {
                clientData.Value.Messages.Clear();
            }

            var messages = _folder.GetMessages().ToArray();
            messages = messages.OrderBy(message => message.ReceivedOn).ToArray();
            logger.LogSampledMetric("MessagesWaitingCount", messages.Length, "count");

            int clientMessageCount = 0;
            foreach (var message in messages)
            {
                double messageAge = (DateTime.UtcNow - message.ReceivedOn.ToUniversalTime()).TotalMinutes;
                logger.LogSampledMetric("MessageAge", messageAge, "minutes", message.Subject);

                bool messageAssigned = false;
                foreach (var clientData in _clients)
                {
                    if (clientData.Value.Evaluator(message))
                    {
                        Logger.InfoFormat("Adding message to client queue: ConversationIndex: {0}, Client ID: {1}, Subject: '{2}'"
                            , message.ConversationIndex, clientData.Key, message.Subject);
                        clientData.Value.Messages.Add(message);
                        messageAssigned = true;
                        clientMessageCount++;
                    }
                }

                if (!messageAssigned)
                {
                    Logger.Info($"Message doesn't match any client: Subject: '{message.Subject}'");
                    RogueMessages.Add(message);
                }
            }

            logger.LogSampledMetric("AllClientMessageCount", clientMessageCount, "count");
            logger.LogSampledMetric("RogueMessageCount", RogueMessages.Count, "count");

            Logger.Info("Finished processing inbox for RecipientsMailboxManagerRouter");
        }

        private struct ClientData
        {
            public MessageEvaluator Evaluator { get; set; }
            public List<IIncomingEmailMessage> Messages { get; set; }
        }


        private readonly IMailFolder _folder;
        private readonly Dictionary<int, ClientData> _clients = new Dictionary<int, ClientData>();
        private int _nextId = 100;
        public List<IIncomingEmailMessage> RogueMessages = new List<IIncomingEmailMessage>();

        private static readonly ILog Logger = LogManager.GetLogger(typeof(RecipientsMailboxManagerRouter));
    }
}
