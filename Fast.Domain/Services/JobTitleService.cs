using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class JobTitleService : ServiceBase<JobTitle>, IJobTitleService
	{
		public JobTitleService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
