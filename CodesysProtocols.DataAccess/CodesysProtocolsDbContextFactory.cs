using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;

namespace CodesysProtocols.DataAccess;

public class CodesysProtocolsDbContextFactory : IDesignTimeDbContextFactory<CodesysProtocolsDbContext>
{
    public CodesysProtocolsDbContext CreateDbContext(string[] args)
    {
        var optionsBuilder = new DbContextOptionsBuilder<CodesysProtocolsDbContext>();
        optionsBuilder.UseSqlServer("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=CodesysProtocols;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False;Command Timeout=30");
        return new CodesysProtocolsDbContext(optionsBuilder.Options);
    }
}