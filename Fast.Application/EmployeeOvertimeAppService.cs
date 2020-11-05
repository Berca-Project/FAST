using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class EmployeeOvertimeAppService : AppServiceBase<EmployeeOvertime>, IEmployeeOvertimeAppService
	{        
        public EmployeeOvertimeAppService(IEmployeeOvertimeService serviceBase) : base(serviceBase)
		{            

        }
    }
}