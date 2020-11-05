using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Application
{
    public class PPLPHAppService : AppServiceBase<PPLPH>, IPPLPHAppService
    {
        private readonly IPPLPHService _PPLphService;
        public PPLPHAppService(IPPLPHService PPLphService) : base(PPLphService)
        {
            _PPLphService = PPLphService;
        }
        public PPLPH GetObjectByID(long id)
        {
            PPLPH ppLph = _PPLphService.GetById(id);
            return ppLph == null ? null : ppLph;
        }
    }
}
