using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistSubmitService : ServiceBase<ChecklistSubmit>, IChecklistSubmitService
    {
		public ChecklistSubmitService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
