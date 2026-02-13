using System.Collections.Generic;

namespace AutoOpener.Core.Jobs
{
    public class AutoOpenProfile
    {
        public string Name { get; set; }
        public int RevitVersion { get; set; }
        public string Server { get; set; }
        public string RsnPath { get; set; }
        public List<string> WorksetsByName { get; set; }
        public bool TriggerOnLogon { get; set; }
        public bool TriggerOnUnlock { get; set; }
        public string DailyTime { get; set; }
        public bool KeepAlive { get; set; }
    }
}
