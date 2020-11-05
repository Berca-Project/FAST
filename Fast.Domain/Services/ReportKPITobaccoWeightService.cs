using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPITobaccoWeightService: ServiceBase<ReportKPITobaccoWeight>, IReportKPITobaccoWeightService
    {
        public ReportKPITobaccoWeightService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
