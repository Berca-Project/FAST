using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIDIMService: ServiceBase<ReportKPIDIM>, IReportKPIDIMService
    {
        public ReportKPIDIMService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
