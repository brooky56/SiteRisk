using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Site.Estimation
{
    public class CascadeResult
    {
        private static string _Name;
        private static double _Prob;
        public CascadeResult(string name, double prob)
        {
            _Name = name;
            _Prob = prob;
        }
        public string Name { get; set; }

        public double Prob { get; set; }
    }
}