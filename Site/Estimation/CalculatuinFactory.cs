using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Site.Estimation
{
    public class CalculatuinFactory
    {
        private static Stage _CalculationTask;
        public CalculatuinFactory(Stage calculationTask)
        {
            _CalculationTask = calculationTask;
        }

        public static double IBThreatProbability(int a, int b = 10) 
        {
            return 1 - Math.Pow(Math.E, (a / b)); 
        }

        public static EffectResult CascadeProbabilty0(int NCount, Stage stage) 
        {
            EffectResult effectResult = new EffectResult(stage)
            {
                N = NCount / 10,
                B = 1 - Math.Pow(Math.E, (-NCount / 10 / 1)),
            };
            
            return effectResult;
        }

        public static EffectResult CascadeProbabilty1(int NVal, int NCount, EffectResult effectResult0, Stage stage)
        {
            EffectResult effectResult1 = new EffectResult(stage)
            {
                N = NVal,
                B = effectResult0.N * NVal / NCount
            };

            return effectResult1;
        }

        public static EffectResult CascadeProbabilty2(int NEffect, int NCount, EffectResult effectResult1, Stage stage)
        {
            EffectResult effectResult2 = new EffectResult(stage)
            {
                N = NEffect,
                B = effectResult1.B * effectResult1.N / NCount
            };

            return effectResult2;
        }
        public static EffectResult CascadeProbabilty3(int NVal, int NEffect, int NCount, EffectResult effectResult2, Stage stage)
        {
            EffectResult effectResult3 = new EffectResult(stage)
            {
                N = NEffect,
                B = effectResult2.N * NVal / NCount
            };

            return effectResult3;
        }

    }
}