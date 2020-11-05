using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistApprovalService : ServiceBase<ChecklistApproval>, IChecklistApprovalService
    {
		public ChecklistApprovalService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
