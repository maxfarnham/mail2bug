using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail2Bug.MessageProcessingStrategies
{
    public class IcmNameResolver : INameResolver
    {
        private readonly HashSet<string> _teamList;
        public IcmNameResolver(IEnumerable<string> teamList)
		{
            var normalizedName = from team in teamList select Normalize(team);
            _teamList = new HashSet<string>(normalizedName);
		}

	    public string Resolve(string alias, string name)
        {
            Logger.InfoFormat("Resolving name for alias/name {0}/{1}", alias, name);

            if (IsValidName(name))
            {
                Logger.DebugFormat("Found name '{0}', returning", name);
                return name;
            }

            if (IsValidName(alias))
            {
                Logger.DebugFormat("Found alias '{0}', returning", name);
                return alias;
            }

	        return null;
        }

        private bool IsValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                Logger.WarnFormat("Trying to resolve a null/empty name");
                return false;
            }

            return true;
        }
    
        private static string Normalize(string name)
        {
            return name.ToLower();
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(IcmNameResolver));
    }
}
