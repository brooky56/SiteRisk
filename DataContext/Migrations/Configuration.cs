namespace DataContext.Migrations
{
    using Microsoft.EntityFrameworkCore;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<DataContext.AnalizDBContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;

          
        }

        protected override void Seed(DataContext.AnalizDBContext context)
        {
            //  This method will be called after migrating to the latest version.

            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data.
        }
    }
}
