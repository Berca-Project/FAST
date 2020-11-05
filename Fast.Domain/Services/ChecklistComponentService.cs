using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistComponentService : ServiceBase<ChecklistComponent>, IChecklistComponentService
    {
		public ChecklistComponentService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
