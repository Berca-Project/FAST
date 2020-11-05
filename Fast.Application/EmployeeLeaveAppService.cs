using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class EmployeeLeaveAppService : AppServiceBase<EmployeeLeave>, IEmployeeLeaveAppService
    {        
        public EmployeeLeaveAppService(IEmployeeLeaveService serviceBase) : base(serviceBase)
		{            
        }
    }
}