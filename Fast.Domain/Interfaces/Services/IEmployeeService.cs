using Fast.Domain.Entities;

namespace Fast.Domain.Interfaces.Services
{
	public interface IEmployeeService : IServiceBase<EmployeeProfile>
	{
        EmployeeProfile GetByName(string name, bool asNoTracking = false);
    }
}
