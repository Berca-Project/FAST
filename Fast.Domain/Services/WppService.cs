using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class WppService : ServiceBase<Wpp>, IWppService
	{
		public WppService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
