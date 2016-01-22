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
            var result = CreateCustomInputResult(query);
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
                        yield return CreateSavedResult(serverName, username);
                    }
                }
            }
        }

        private static Result CreateCustomInputResult(Query query)
        {
            if(query.Search == null || query.Search.Trim() == string.Empty)
                return null;

            var serverName = query.FirstSearch;
            var userName = query.SecondSearch;
            var password = query.ThirdSearch;

            var result = new Result(serverName, ICON_PATH, string.Format("User:{0} Password:{1}", userName, password))
                         {
                             Action = context => Run(serverName, userName, password),
                             Score = int.MaxValue
                         };
            return result;
        }

        private static Result CreateSavedResult(string serverName, string userName)
        {
            var result = new Result(serverName, ICON_PATH, userName) {Action = context => Run(serverName)};
            return result;
        }

        private static bool Run(string serverName, string userName = null, string password = null)
        {
            //thanks to https://stackoverflow.com/questions/11296819/run-mstsc-exe-with-specified-username-and-password
            using(var rdcProcess = new Process())
            {
                if(string.IsNullOrEmpty(userName) == false)
                {
                    rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                    var args = string.IsNullOrEmpty(password) ? "/generic:TERMSRV/{0} /user:{1}" : "/generic:TERMSRV/{0} /user:{1} /pass:{2}";
                    rdcProcess.StartInfo.Arguments = string.Format(args, serverName, userName, password);
                    rdcProcess.Start();
                }
                rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                rdcProcess.StartInfo.Arguments = string.Format("/v:{0}", serverName); // ip or name of computer to connect
                rdcProcess.Start();

                //rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                //rdcProcess.StartInfo.Arguments = string.Format("/delete:TERMSRV/{0}", serverName); // delete credentials
                //rdcProcess.Start();
            }
            return true;
        }
    }
}