using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Win32;

namespace Wox.Plugin.RDP
{
    public class RDPPlugin : IPlugin
    {
        private const string REGISTRY_KEY = @"Software\Microsoft\Terminal Server Client\Servers";
        private const string ICON_PATH = @"Images\icon.png";

        public void Init(PluginInitContext context)
        {
        }

        public List<Result> Query(Query query)
        {
            var results = new List<Result>();
            var result = CreateCustomInputResult(query.Search);
            if(result != null)
                results.Add(result);
            results.AddRange(GetRdpSessions(query.Search));
            return results;
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

        private static Result CreateCustomInputResult(string search)
        {
            if(search == null || search.Trim() == string.Empty)
                return null;
            var p = search.IndexOf(' ');
            string serverName;
            string userName = null;
            if(p < 0)
                serverName = search;
            else
            {
                serverName = search.Substring(0, p);
                userName = search.Substring(p + 1);
            }
            var reult = CreateResult(serverName, userName);
            reult.Score = int.MaxValue;
            return reult;
        }

        private static Result CreateResult(string serverName, string userName)
        {
            var result = new Result(serverName, ICON_PATH, userName) {Action = context => Run(serverName, userName)};
            return result;
        }

        private static bool Run(string serverName, string userName)
        {
            //thanks to https://stackoverflow.com/questions/11296819/run-mstsc-exe-with-specified-username-and-password
            var rdcProcess = new Process();
            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
            rdcProcess.StartInfo.Arguments = string.Format("/generic:TERMSRV/{0} /user:{1}", serverName, userName);
            rdcProcess.Start();

            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
            rdcProcess.StartInfo.Arguments = string.Format("/v:{0}", serverName); // ip or name of computer to connect
            rdcProcess.Start();

            //rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
            //rdcProcess.StartInfo.Arguments = string.Format("/delete:TERMSRV/{0}", serverName); // delete credentials
            //rdcProcess.Start();
            return true;
        }
    }
}