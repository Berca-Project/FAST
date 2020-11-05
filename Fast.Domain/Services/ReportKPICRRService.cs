using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPICRRService: ServiceBase<ReportKPICRR>, IReportKPICRRService
    {
        public ReportKPICRRService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
