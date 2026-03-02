using System;
using System.Collections.Generic;

namespace AutoOpener.Core.Jobs
{
    public class JobResult
    {
        public Guid JobId { get; set; }
        public int RevitVersion { get; set; }
        public bool Succeeded { get; set; }
        public string Message { get; set; }
        public string JobType { get; set; }
        public List<string> OpenedModelPaths { get; set; }
        public DateTime CompletedUtc { get; set; }
        public JobResult()
        {
            OpenedModelPaths = new List<string>();
        }
    }
}
