using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class MachineAllocationService : ServiceBase<MachineAllocation>, IMachineAllocationService
	{
		public MachineAllocationService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
