using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistSubmitAppService : AppServiceBase<ChecklistSubmit>, IChecklistSubmitAppService
    {
        public ChecklistSubmitAppService(IChecklistSubmitService serviceBase) : base(serviceBase)
        {
        }
    }
}