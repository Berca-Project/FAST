using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class InputTargetService: ServiceBase<InputTarget>, IInputTargetService
    {
        public InputTargetService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
