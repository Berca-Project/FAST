using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIDustService: ServiceBase<ReportKPIDust>, IReportKPIDustService
    {
        public ReportKPIDustService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
