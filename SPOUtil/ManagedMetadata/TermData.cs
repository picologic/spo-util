using System.Collections.Generic;

namespace SPOUtil.ManagedMetadata
{
    public class TermData
    {
        public string Term { get; set; }
        public IDictionary<string, string> Properties { get; set; }

        public TermData() {
            Properties = new Dictionary<string, string>();
        }
    }
}
