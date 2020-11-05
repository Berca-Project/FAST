using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportRemarksService: ServiceBase<ReportRemarks>, IReportRemarksService
    {
        public ReportRemarksService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
