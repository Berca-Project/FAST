using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MealRequestAppService : AppServiceBase<MealRequest>, IMealRequestAppService
	{
		public MealRequestAppService(IMealRequestService serviceBase) : base(serviceBase)
		{
		}
	}
}