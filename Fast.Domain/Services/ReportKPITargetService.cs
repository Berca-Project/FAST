using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPITargetService: ServiceBase<ReportKPITarget>, IReportKPITargetService
    {
        public ReportKPITargetService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
