using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIStickPerPackService: ServiceBase<ReportKPIStickPerPack>, IReportKPIStickPerPackService
    {
        public ReportKPIStickPerPackService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
