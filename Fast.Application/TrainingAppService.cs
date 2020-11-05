using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class TrainingAppService: AppServiceBase<Training>, ITrainingAppService
    {
        public TrainingAppService(ITrainingService serviceBase) : base(serviceBase)
        {
        }
    }
}
