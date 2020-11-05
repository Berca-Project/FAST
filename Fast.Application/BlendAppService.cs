using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class BlendAppService : AppServiceBase<Blend>, IBlendAppService
	{
		public BlendAppService(IBlendService serviceBase) : base(serviceBase)
		{
		}
	}
}