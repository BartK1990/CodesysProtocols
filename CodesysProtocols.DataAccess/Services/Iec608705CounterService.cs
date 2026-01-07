using CodesysProtocols.DataAccess.Entities;
using CodesysProtocols.DataAccess.Enums;
using Microsoft.EntityFrameworkCore;

namespace CodesysProtocols.DataAccess.Services;

public class Iec608705CounterService : IIec608705CounterService
{
    private readonly CodesysProtocolsDbContext _context;

    public Iec608705CounterService(CodesysProtocolsDbContext context)
    {
        _context = context;
    }

    public async Task<long> GetCounterValueAsync(Codesys23Iec608705CounterType counterType)
    {
        var counter = await _context.Codesys23Iec608705Counters.SingleOrDefaultAsync(c => c.Name == counterType.ToString());
        return counter?.Value ?? 0;
    }

    public async Task IncrementCounterValueAsync(Codesys23Iec608705CounterType counterType)
    {
        var updated = await _context.Codesys23Iec608705Counters
            .Where(c => c.Name == counterType.ToString())
            .ExecuteUpdateAsync(s => s.SetProperty(c => c.Value, c => c.Value + 1));
        if (updated == 0)
        {
            _context.Codesys23Iec608705Counters.Add(new Codesys23Iec608705Counter
            {
                Name = counterType.ToString(),
                Value = 1,
            });

            await _context.SaveChangesAsync();
        }
    }
}