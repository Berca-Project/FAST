using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using Fast.Infra.CrossCutting.Common;

namespace Fast.Application
{
	public class EmployeeAppService : AppServiceBase<EmployeeProfile>, IEmployeeAppService
	{
        private readonly IEmployeeService _empService;
        public EmployeeAppService(IEmployeeService serviceBase) : base(serviceBase)
		{
            _empService = serviceBase;

        }

        public string GetNameById(long id)
        {
            EmployeeProfile emp = _empService.GetById(id);
            return emp == null ? string.Empty : emp.FullName;
        }
        public string GetEmployeeIdByID(long id)
        {
            EmployeeProfile emp = _empService.GetById(id);
            return emp == null ? string.Empty : emp.EmployeeID;
        }
        public string GetByName(string name, bool asNoTracking = false)
        {
            EmployeeProfile emp = _empService.GetByName(name, asNoTracking);
            string result = emp != null ? JsonHelper<EmployeeProfile>.Serialize(emp) : string.Empty;

            return result;
        }
    }
}