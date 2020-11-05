using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIProdVolService: ServiceBase<ReportKPIProdVol>, IReportKPIProdVolService
    {
        public ReportKPIProdVolService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
