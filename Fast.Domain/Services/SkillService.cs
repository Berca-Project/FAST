using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class SkillService : ServiceBase<Skill>, ISkillService
	{
		public SkillService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
