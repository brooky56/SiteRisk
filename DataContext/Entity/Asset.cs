using System;
using System.Collections.Generic;

namespace DataContext
{
    public class Asset
    {
        public Guid AssetId { get; set; }
        public string Name { get; set; }
        public string BName { get; set; }

        public Threat Threat { get; set; }

        public Guid ThreatId { get; set; }

        public ICollection<Threat> Threats { get; set; }
    }
}