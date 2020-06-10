using MySql.Data.EntityFramework;
using System;
using System.Data.Entity;

namespace DataContext
{
    [DbConfigurationType(typeof(MySqlEFConfiguration))]
    public class AnalizDBContext : DbContext
    {
        public DbSet<Asset> Assets { get; set; }
        public DbSet<Consequence> Consequences { get; set; }
        public DbSet<Threat> Threats { get; set; }
        public DbSet<Vulnerability> Vulnerabilities { get; set; }

        public AnalizDBContext() : base("DefaultConnection")
        {

        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Threat>().HasKey(e => e.ThreatId);
            modelBuilder.Entity<Threat>().Property(e => e.Name).IsRequired();
            modelBuilder.Entity<Threat>().Property(e => e.BName).IsRequired();
            modelBuilder.Entity<Threat>().HasRequired<Asset>(t => t.Asset)
                .WithMany(a => a.Threats)
                .HasForeignKey<Guid>(t => t.AssetId);
            modelBuilder.Entity<Threat>().HasRequired<Vulnerability>(t => t.Vulnerability)
                .WithMany(v => v.Threats)
                .HasForeignKey<Guid>(t => t.VulnerabilityId);

            modelBuilder.Entity<Asset>().HasKey(a => a.AssetId);
            modelBuilder.Entity<Asset>().Property(a => a.Name).IsRequired();
            modelBuilder.Entity<Asset>().Property(a => a.BName).IsRequired();
            modelBuilder.Entity<Asset>()
                .HasMany<Threat>(a => a.Threats)
                .WithRequired(t => t.Asset)
                .HasForeignKey<Guid>(t => t.AssetId);

            modelBuilder.Entity<Vulnerability>().HasKey(v => v.VulnerabilityId);
            modelBuilder.Entity<Vulnerability>().Property(v => v.Name).IsRequired();
            modelBuilder.Entity<Vulnerability>().Property(v => v.BName).IsRequired();
            modelBuilder.Entity<Vulnerability>()
                .HasMany<Threat>(v => v.Threats)
                .WithRequired(t => t.Vulnerability)
                .HasForeignKey<Guid>(t => t.VulnerabilityId);

           
        }
    }
}