using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class SkillAppService : AppServiceBase<Skill>, ISkillAppService
	{
		public SkillAppService(ISkillService serviceBase) : base(serviceBase)
		{
		}
	}
}