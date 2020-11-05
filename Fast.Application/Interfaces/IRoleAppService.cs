namespace Fast.Application.Interfaces
{
	public interface IRoleAppService : IAppServiceBase
	{
        void RemoveEntity(string roleName);
		string GetNameById(long id);
		string GetByName(string name, bool asNoTracking = false);
	}
}