using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class EmployeeOvertimeService : ServiceBase<EmployeeOvertime>, IEmployeeOvertimeService
	{
		public EmployeeOvertimeService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
