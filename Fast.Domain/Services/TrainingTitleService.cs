using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class TrainingTitleService : ServiceBase<TrainingTitle>, ITrainingTitleService
	{
		public TrainingTitleService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
