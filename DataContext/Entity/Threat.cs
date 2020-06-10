using System;
using System.Collections.Generic;

namespace DataContext
{
    public class Threat
    {
        public Guid ThreatId { get; set; }
        public string Name { get; set; }
        public string BName { get; set; }

        public Guid AssetId { get; set; }
        public Asset Asset { get; set; }

        public Guid VulnerabilityId { get; set; }
        public Vulnerability Vulnerability { get; set; }

    }
}