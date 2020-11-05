using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class TrainingTitleMachineTypeAppService : AppServiceBase<TrainingMachineType>, ITrainingTitleMachineTypeAppService
	{
		public TrainingTitleMachineTypeAppService(ITrainingTitleMachineTypeService serviceBase) : base(serviceBase)
		{
		}
	}
}
