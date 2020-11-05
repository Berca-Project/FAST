using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class WppStpService : ServiceBase<WppStp>, IWppStpService
	{
		public WppStpService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
