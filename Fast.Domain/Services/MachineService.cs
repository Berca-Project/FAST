using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class MachineService : ServiceBase<Machine>, IMachineService
	{
		public MachineService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
