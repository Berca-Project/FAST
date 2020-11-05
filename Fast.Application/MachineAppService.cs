using Fast.Application;
using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MachineAppService : AppServiceBase<Machine>, IMachineAppService
	{
		public MachineAppService(IMachineService serviceBase) : base(serviceBase)
		{
		}
	}
}