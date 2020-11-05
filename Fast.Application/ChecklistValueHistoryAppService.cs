using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistValueHistoryAppService : AppServiceBase<ChecklistValueHistory>, IChecklistValueHistoryAppService
    {
        public ChecklistValueHistoryAppService(IChecklistValueHistoryService serviceBase) : base(serviceBase)
        {
        }
    }
}