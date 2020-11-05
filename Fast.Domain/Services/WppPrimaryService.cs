using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class WppPrimaryService : ServiceBase<WppPrimary>, IWppPrimaryService
	{
		public WppPrimaryService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
