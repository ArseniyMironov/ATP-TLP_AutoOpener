using System.Collections.Generic;

namespace AutoOpener.Core.Models
{
    public class Profile
    {
        public string Name { get; set; }
        public int RevitVersion { get; set; }
        public bool OpenReadOnly { get; set; }
        public List<ModelTask> Models { get; set; }

        public Profile()
        {
            Models = new List<ModelTask>();
        }
    }
}
