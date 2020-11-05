using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class MppService : ServiceBase<Mpp>, IMppService
	{
		public MppService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
