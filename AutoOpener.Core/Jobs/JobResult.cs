using System;

namespace AutoOpener.Core.Jobs
{
    public class JobResult
    {
        public Guid JobId { get; set; }
        public bool Succeeded { get; set; }
        public string Message { get; set; }
        public string OutputFile { get; set; }
        public string JobType { get; set; }
        public string ModelPath { get; set; }
        public DateTime CompletedUtc { get; set; }
    }
}
