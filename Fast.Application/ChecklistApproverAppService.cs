using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistApproverAppService : AppServiceBase<ChecklistApprover>, IChecklistApproverAppService
    {
        public ChecklistApproverAppService(IChecklistApproverService serviceBase) : base(serviceBase)
        {
        }
    }
}