namespace Fast.Application.Interfaces
{
	public interface IJobTitleAppService : IAppServiceBase
	{
		string GetRoleNameByJobTitleId(long id);
	}
}