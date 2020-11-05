using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ReferenceDetailService : ServiceBase<ReferenceDetail>, IReferenceDetailService
	{
		public ReferenceDetailService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
