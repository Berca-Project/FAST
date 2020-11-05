using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class BrandConversionService : ServiceBase<BrandConversion>, IBrandConversionService
	{
		public BrandConversionService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
