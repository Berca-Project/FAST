using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MachineAllocationAppService : AppServiceBase<MachineAllocation>, IMachineAllocationAppService
	{
		public MachineAllocationAppService(IMachineAllocationService serviceBase) : base(serviceBase)
		{
		}
	}
}