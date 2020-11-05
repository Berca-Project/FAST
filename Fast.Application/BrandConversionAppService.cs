using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class BrandConversionAppService : AppServiceBase<BrandConversion>, IBrandConversionAppService
	{
		public BrandConversionAppService(IBrandConversionService serviceBase) : base(serviceBase)
		{
		}
	}
}