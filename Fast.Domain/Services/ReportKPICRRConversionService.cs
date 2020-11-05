using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPICRRConversionService: ServiceBase<ReportKPICRRConversion>, IReportKPICRRConversionService
    {
        public ReportKPICRRConversionService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
