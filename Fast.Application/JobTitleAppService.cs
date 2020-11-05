using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class JobTitleAppService : AppServiceBase<JobTitle>, IJobTitleAppService
	{
		public readonly IJobTitleService _serviceBase;
		public JobTitleAppService(IJobTitleService serviceBase) : base(serviceBase)
		{
			_serviceBase = serviceBase;
		}

		public string GetRoleNameByJobTitleId(long id)
		{
			JobTitle jt = _serviceBase.GetById(id);
			return jt == null ? string.Empty : jt.RoleName;
		}
	}
}