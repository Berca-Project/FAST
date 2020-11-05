using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistApprovalAppService : AppServiceBase<ChecklistApproval>, IChecklistApprovalAppService
    {
        public ChecklistApprovalAppService(IChecklistApprovalService serviceBase) : base(serviceBase)
        {
        }
    }
}