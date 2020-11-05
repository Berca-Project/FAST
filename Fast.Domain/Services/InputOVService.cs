using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class InputOVService: ServiceBase<InputOV>, IInputOVService
    {
        public InputOVService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
