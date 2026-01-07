using CodesysProtocols.DataAccess.Entities;
using Microsoft.EntityFrameworkCore;

namespace CodesysProtocols.DataAccess;

public class CodesysProtocolsDbContext : DbContext
{
    public CodesysProtocolsDbContext(DbContextOptions<CodesysProtocolsDbContext> options)
        : base(options)
    {
    }

    public DbSet<Codesys23Iec608705Counter> Codesys23Iec608705Counters => Set<Codesys23Iec608705Counter>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<Codesys23Iec608705Counter>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.Name).IsRequired();
            entity.Property(e => e.Value).IsRequired();
        });
    }
}