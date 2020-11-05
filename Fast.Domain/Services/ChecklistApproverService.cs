using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistApproverService : ServiceBase<ChecklistApprover>, IChecklistApproverService
    {
		public ChecklistApproverService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
