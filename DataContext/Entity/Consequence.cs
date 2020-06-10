using System;
using System.Collections.Generic;

namespace DataContext
{
    public class Consequence
    {
        public Guid ConsequenceId { get; set; }
        public string Name { get; set; }
        public string BName { get; set; }

        public virtual ICollection<Threat> Threats { get; set; }
        public virtual  ICollection<Asset> Assets { get; set; }

        public Consequence()
        {
            Threats = new List<Threat>();
            Assets = new List<Asset>();
        }
    }
}