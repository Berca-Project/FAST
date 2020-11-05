using Fast.Domain.Entities;

namespace Fast.Domain.Interfaces.Services
{
	public interface IRoleService : IServiceBase<Role>
	{
		Role GetByName(string name, bool asNoTracking = false);
	}
}
