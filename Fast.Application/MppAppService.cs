using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MppAppService : AppServiceBase<Mpp>, IMppAppService
	{
		public MppAppService(IMppService serviceBase) : base(serviceBase)
		{
		}
	}
}