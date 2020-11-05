using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIYieldService: ServiceBase<ReportKPIYield>, IReportKPIYieldService
    {
        public ReportKPIYieldService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
