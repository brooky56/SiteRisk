using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataContext
{
    public class EntitySeed : DropCreateDatabaseIfModelChanges<AnalizDBContext>
    {
        protected override void Seed(AnalizDBContext context)
        {
            List<Threat> threats = new List<Threat>
        {
            new Threat {
                ThreatId = Guid.NewGuid(),
                Name = "УБИ. 001",
                BName = "Угроза автоматического распространения вредоносного кода в грид-системе"
                },
            new Threat {
                ThreatId = Guid.NewGuid(),
                Name = "УБИ. 002",
                BName = "Угроза агрегирования данных, передаваемых в грид-системе"
            },
            new Threat {
                ThreatId = Guid.NewGuid(),
                Name = "УБИ. 003",
                BName = "Угроза анализа криптографических алгоритмов и их реализации"
            },
        };

            foreach (Threat item in threats)
            {
                context.Threats.Add(item);
            }

            context.SaveChanges();

        }
    }
}
