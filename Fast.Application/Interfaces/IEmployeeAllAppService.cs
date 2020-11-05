namespace Fast.Application.Interfaces
{
	public interface IEmployeeAllAppService : IAppServiceBase
	{
        string GetNameById(long id);
        string GetEmployeeIdByID(long id);
        string GetByName(string name, bool asNoTracking = false);
    }
}