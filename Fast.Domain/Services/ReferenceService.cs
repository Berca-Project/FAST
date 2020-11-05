using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ReferenceService : ServiceBase<Reference>, IReferenceService
	{
		public ReferenceService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
