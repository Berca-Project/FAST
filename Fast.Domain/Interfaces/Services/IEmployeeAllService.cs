using Fast.Domain.Entities;

namespace Fast.Domain.Interfaces.Services
{
	public interface IEmployeeAllService : IServiceBase<EmployeeProfileAll>
	{
        EmployeeProfileAll GetByName(string name, bool asNoTracking = false);
    }
}
