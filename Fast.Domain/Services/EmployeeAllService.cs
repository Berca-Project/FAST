using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class EmployeeAllService : ServiceBase<EmployeeProfileAll>, IEmployeeAllService
    {
        private readonly IUnitOfWork _unitOfWork;
        private readonly IRepositoryBase<EmployeeProfileAll> _repository;

        public EmployeeAllService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
            _unitOfWork = unitOfWork;
            _repository = _unitOfWork.Repository<EmployeeProfileAll>();
        }

        public EmployeeProfileAll GetByName(string name, bool asNoTracking = false)
        {
            return _repository.GetByName(name, asNoTracking);
        }
    }
}
