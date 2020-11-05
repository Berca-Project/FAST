using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class BrandService : ServiceBase<Brand>, IBrandService
	{
		public BrandService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
