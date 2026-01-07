using CodesysProtocols.DataAccess.Enums;

namespace CodesysProtocols.DataAccess.Services;

public interface IIec608705CounterService
{
    Task<long> GetCounterValueAsync(Codesys23Iec608705CounterType counterType);

    Task IncrementCounterValueAsync(Codesys23Iec608705CounterType counterType);
}