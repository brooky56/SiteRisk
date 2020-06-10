namespace DataContext.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class First : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Assets",
                c => new
                    {
                        AssetId = c.Guid(nullable: false),
                        Name = c.String(nullable: false, unicode: false),
                        BName = c.String(nullable: false, unicode: false),
                        ThreatId = c.Guid(nullable: false),
                        Threat_ThreatId = c.Guid(),
                        Consequence_ConsequenceId = c.Guid(),
                    })
                .PrimaryKey(t => t.AssetId)
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId)
                .ForeignKey("dbo.Consequences", t => t.Consequence_ConsequenceId)
                .Index(t => t.Threat_ThreatId)
                .Index(t => t.Consequence_ConsequenceId);
            
            CreateTable(
                "dbo.Threats",
                c => new
                    {
                        ThreatId = c.Guid(nullable: false),
                        Name = c.String(nullable: false, unicode: false),
                        BName = c.String(nullable: false, unicode: false),
                        AssetId = c.Guid(nullable: false),
                        VulnerabilityId = c.Guid(nullable: false),
                        Consequence_ConsequenceId = c.Guid(),
                    })
                .PrimaryKey(t => t.ThreatId)
                .ForeignKey("dbo.Vulnerabilities", t => t.VulnerabilityId, cascadeDelete: true)
                .ForeignKey("dbo.Assets", t => t.AssetId, cascadeDelete: true)
                .ForeignKey("dbo.Consequences", t => t.Consequence_ConsequenceId)
                .Index(t => t.AssetId)
                .Index(t => t.VulnerabilityId)
                .Index(t => t.Consequence_ConsequenceId);
            
            CreateTable(
                "dbo.Vulnerabilities",
                c => new
                    {
                        VulnerabilityId = c.Guid(nullable: false),
                        Name = c.String(nullable: false, unicode: false),
                        BName = c.String(nullable: false, unicode: false),
                        ThreatId = c.Guid(nullable: false),
                        Threat_ThreatId = c.Guid(),
                    })
                .PrimaryKey(t => t.VulnerabilityId)
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId)
                .Index(t => t.Threat_ThreatId);
            
            CreateTable(
                "dbo.Consequences",
                c => new
                    {
                        ConsequenceId = c.Guid(nullable: false),
                        Name = c.String(unicode: false),
                        BName = c.String(unicode: false),
                    })
                .PrimaryKey(t => t.ConsequenceId);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Threats", "Consequence_ConsequenceId", "dbo.Consequences");
            DropForeignKey("dbo.Assets", "Consequence_ConsequenceId", "dbo.Consequences");
            DropForeignKey("dbo.Threats", "AssetId", "dbo.Assets");
            DropForeignKey("dbo.Assets", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.Threats", "VulnerabilityId", "dbo.Vulnerabilities");
            DropForeignKey("dbo.Vulnerabilities", "Threat_ThreatId", "dbo.Threats");
            DropIndex("dbo.Vulnerabilities", new[] { "Threat_ThreatId" });
            DropIndex("dbo.Threats", new[] { "Consequence_ConsequenceId" });
            DropIndex("dbo.Threats", new[] { "VulnerabilityId" });
            DropIndex("dbo.Threats", new[] { "AssetId" });
            DropIndex("dbo.Assets", new[] { "Consequence_ConsequenceId" });
            DropIndex("dbo.Assets", new[] { "Threat_ThreatId" });
            DropTable("dbo.Consequences");
            DropTable("dbo.Vulnerabilities");
            DropTable("dbo.Threats");
            DropTable("dbo.Assets");
        }
    }
}
