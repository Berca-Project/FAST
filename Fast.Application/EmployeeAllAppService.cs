using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using Fast.Infra.CrossCutting.Common;

namespace Fast.Application
{
	public class EmployeeAllAppService : AppServiceBase<EmployeeProfileAll>, IEmployeeAllAppService
	{
        private readonly IEmployeeAllService _empService;
        public EmployeeAllAppService(IEmployeeAllService serviceBase) : base(serviceBase)
		{
            _empService = serviceBase;
        }

        public string GetNameById(long id)
        {
            EmployeeProfileAll emp = _empService.GetById(id);
            return emp == null ? string.Empty : emp.FullName;
        }
        public string GetEmployeeIdByID(long id)
        {
			EmployeeProfileAll emp = _empService.GetById(id);
            return emp == null ? string.Empty : emp.EmployeeID;
        }
        public string GetByName(string name, bool asNoTracking = false)
        {
			EmployeeProfileAll emp = _empService.GetByName(name, asNoTracking);
            string result = emp != null ? JsonHelper<EmployeeProfileAll>.Serialize(emp) : string.Empty;

            return result;
        }
    }
}