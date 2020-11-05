using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class MealRequestService : ServiceBase<MealRequest>, IMealRequestService
	{
		public MealRequestService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
