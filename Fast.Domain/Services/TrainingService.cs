using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class TrainingService:ServiceBase<Training>, ITrainingService
    {
        public TrainingService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
