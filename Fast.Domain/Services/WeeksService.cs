using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class WeeksService:ServiceBase<Weeks>, IWeeksService
    {
        public WeeksService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
