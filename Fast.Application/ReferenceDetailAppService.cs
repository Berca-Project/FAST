using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class ReferenceDetailAppService : AppServiceBase<ReferenceDetail>, IReferenceDetailAppService
	{
		public ReferenceDetailAppService(IReferenceDetailService serviceBase) : base(serviceBase)
		{
		}
	}
}