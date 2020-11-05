using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ManPowerService : ServiceBase<ManPower>, IManPowerService
	{
		public ManPowerService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
