using System.Text;
using log4net;
using Mail2Bug.Email.EWS;

namespace Mail2Bug.Email
{
    internal class AckEmailHandler
    {
        private readonly Config.InstanceConfig _config;

        public AckEmailHandler(Config.InstanceConfig config)
        {
            _config = config;
        }

        /// <summary>
        /// Send mail announcing receipt of new ticket
        /// </summary>
        public void SendAckEmail(IIncomingEmailMessage originalMessage, string workItemId)
        {
            // Don't send ack emails if it's disabled in configuration
            if (!_config.EmailSettings.SendAckEmails)
            {
                Logger.DebugFormat("Ack emails disabled in configuration - skipping");
                return;
            }

            var ewsMessage = originalMessage as EWSIncomingMessage;
            if (ewsMessage != null)
            {
                HandleEWSMessage(ewsMessage, workItemId);
            }
        }

        private void HandleEWSMessage(EWSIncomingMessage originalMessage, string workItemId)
        {
            originalMessage.Reply(GetReplyContents(workItemId), _config.EmailSettings.AckEmailsRecipientsAll);
        }

        private string GetReplyContents(string workItemId)
        {
            var bodyBuilder = new StringBuilder();
            bodyBuilder.Append(_config.EmailSettings.GetReplyTemplate());

            if (_config.TfsServerConfig != null)
            {
                bodyBuilder.Replace("[BUGID]", workItemId);
                bodyBuilder.Replace("[TFSCollectionUri]", _config.TfsServerConfig.CollectionUri);
            }
            
            if (_config.IcmClientConfig != null)
            {
                bodyBuilder.Replace("[IncidentId]", workItemId);
                bodyBuilder.Replace("[TFSCollectionUri]", _config.IcmClientConfig.IcmUri);
            }
 
            return bodyBuilder.ToString();
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(AckEmailHandler));
    }
}
