namespace Fast.Application.Interfaces
{
	public interface IEmployeeAppService : IAppServiceBase
	{
        string GetNameById(long id);
        string GetEmployeeIdByID(long id);
        string GetByName(string name, bool asNoTracking = false);
    }
}