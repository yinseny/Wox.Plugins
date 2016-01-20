using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Win32;

namespace Wox.Plugin.RDP
{
    public class RDPPlugin : IPlugin
    {
        private const string REGISTRY_KEY = @"Software\Microsoft\Terminal Server Client\Servers";

        public void Init(PluginInitContext context)
        {
        }

        public List<Result> Query(Query query)
        {
            var sessions = GetRdpSessions(query.Search);
            return sessions.ToList();
        }

        private static IEnumerable<Result> GetRdpSessions(string search)
        {
            using(RegistryKey rdpSessions = Registry.CurrentUser.OpenSubKey(REGISTRY_KEY))
            {
                if(rdpSessions == null)
                    yield break;
                foreach(var serverName in rdpSessions.GetSubKeyNames())
                {
                    using(var serverKey = rdpSessions.OpenSubKey(serverName))
                    {
                        if(serverKey == null)
                            continue;
                        if(serverName.StartsWith(search) == false)
                            continue;
                        var username = serverKey.GetValue("UsernameHint", string.Empty).ToString();
                        yield return CreateResult(serverName, username);
                    }
                }
            }
        }

        private static Result CreateResult(string serverName, string userName)
        {
            var result = new Result(serverName, @"Images\icon.png", userName) {Action = context => Run(serverName)};
            return result;
        }

        private static bool Run(string serverName)
        {
            Process.Start("mstsc.exe", string.Format("/v:{0}", serverName));
            return true;
        }
    }
}