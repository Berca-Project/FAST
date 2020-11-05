using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistValueHistoryService : ServiceBase<ChecklistValueHistory>, IChecklistValueHistoryService
    {
		public ChecklistValueHistoryService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
