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
    public class PPLPHExtrasAppService : AppServiceBase<PPLPHExtras>, IPPLPHExtrasAppService
    {
        private readonly IPPLPHExtrasService _PPLphExtrasService;
        public PPLPHExtrasAppService(IPPLPHExtrasService PPLphExtrasService) : base(PPLphExtrasService)
        {
            _PPLphExtrasService = PPLphExtrasService;
        }
    }
}
