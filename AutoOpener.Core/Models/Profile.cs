using System.Collections.Generic;

namespace AutoOpener.Core.Models
{
    public class Profile
    {
        public string Name { get; set; }
        public int RevitVersion { get; set; }
        public string ModelPathOrRsn { get; set; }
        public List<string> WorksetsByName { get; set; }
        public bool OpenReadOnly { get; set; }
    }
}
