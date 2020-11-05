using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class EmployeeLeaveService : ServiceBase<EmployeeLeave>, IEmployeeLeaveService
	{
		public EmployeeLeaveService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
