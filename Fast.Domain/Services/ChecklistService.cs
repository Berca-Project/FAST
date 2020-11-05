using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistService : ServiceBase<Checklist>, IChecklistService
    {
		public ChecklistService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
