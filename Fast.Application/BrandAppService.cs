using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class BrandAppService : AppServiceBase<Brand>, IBrandAppService
	{
		public BrandAppService(IBrandService serviceBase) : base(serviceBase)
		{
		}
	}
}