using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MaterialCodeAppService : AppServiceBase<MaterialCode>, IMaterialCodeAppService
	{
		public MaterialCodeAppService(IMaterialCodeService serviceBase) : base(serviceBase)
		{
		}
	}
}