using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class EmployeeService : ServiceBase<EmployeeProfile>, IEmployeeService
    {
        private readonly IUnitOfWork _unitOfWork;
        private readonly IRepositoryBase<EmployeeProfile> _repository;

        public EmployeeService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
            _unitOfWork = unitOfWork;
            _repository = _unitOfWork.Repository<EmployeeProfile>();
        }

        public EmployeeProfile GetByName(string name, bool asNoTracking = false)
        {
            return _repository.GetByName(name, asNoTracking);
        }
    }
}
