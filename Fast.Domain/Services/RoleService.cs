using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class RoleService : ServiceBase<Role>, IRoleService
	{
		private readonly IUnitOfWork _unitOfWork;
		private readonly IRepositoryBase<Role> _repository;

		public RoleService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
			_unitOfWork = unitOfWork;
			_repository = _unitOfWork.Repository<Role>();
		}

		public Role GetByName(string name, bool asNoTracking = false)
		{
			return _repository.GetByName(name, asNoTracking);
		}
	}
}
