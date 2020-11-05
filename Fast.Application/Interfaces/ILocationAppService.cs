namespace Fast.Application.Interfaces
{
	public interface ILocationAppService : IAppServiceBase
	{
		string GetLocationFullCode(long locationID);
        long GetLocationID(string locationCode);
    }
}