using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class BlendService : ServiceBase<Blend>, IBlendService
	{
		public BlendService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
