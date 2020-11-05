using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ChecklistValueService : ServiceBase<ChecklistValue>, IChecklistValueService
    {
		public ChecklistValueService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
