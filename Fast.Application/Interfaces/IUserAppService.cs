namespace Fast.Application.Interfaces
{
	public interface IUserAppService : IAppServiceBase
	{
        string GetNameById(long id);
    }
}