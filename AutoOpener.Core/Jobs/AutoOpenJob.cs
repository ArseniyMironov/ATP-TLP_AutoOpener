using System;
using System.Collections.Generic;

namespace AutoOpener.Core.Jobs
{
    public class AutoOpenJob
    {
        public Guid Id { get; set; }
        public int RevitVersion { get; set; }
        public string RsnPath { get; set; }
        public List<string> WorksetsByName { get; set; }
        public bool CreateNewLocal { get; set; } = true;
        public bool OpenReadOnly { get; set; } = false;
        public DateTime RequestedAtUtc { get; set; } = DateTime.UtcNow;
    }
}
