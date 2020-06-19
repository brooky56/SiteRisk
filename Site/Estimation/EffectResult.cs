using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Site.Estimation
{
    public class EffectResult
    {
        private static Stage _CurrentStage;
        public EffectResult(Stage stage)
        {
            _CurrentStage = stage;
        }

        public double N;

        public double B;

        public double NValNext()
        {
            return _CurrentStage.N;
        }

        public double BValPrev()
        {
            return _CurrentStage.B;
        }
    }
}