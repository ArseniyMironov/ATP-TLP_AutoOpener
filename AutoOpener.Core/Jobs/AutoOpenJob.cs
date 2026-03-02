using AutoOpener.Core.Models;
using System;
using System.Collections.Generic;

namespace AutoOpener.Core.Jobs
{
    public class AutoOpenJob
    {
        public Guid Id { get; set; }
        public int RevitVersion { get; set; }
        public bool CreateNewLocal { get; set; } = true;
        public List<ModelTask> Models { get; set; }
        public AutoOpenJob()
        {
            Models = new List<ModelTask>();
        }
    }
}
