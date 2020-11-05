using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIRipperInfoService: ServiceBase<ReportKPIRipperInfo>, IReportKPIRipperInfoService
    {
        public ReportKPIRipperInfoService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
