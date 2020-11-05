using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class TrainingTitleMachineTypeService : ServiceBase<TrainingMachineType>, ITrainingTitleMachineTypeService
	{
		public TrainingTitleMachineTypeService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
