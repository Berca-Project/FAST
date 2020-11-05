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
    public class LPHValueHistoriesAppService: AppServiceBase<LPHValueHistories>, ILPHValueHistoriesAppService
    {
        private readonly ILPHValueHistoriesService _lphValueHistoriesService;
        public LPHValueHistoriesAppService(ILPHValueHistoriesService lphValueHistoriesService) : base(lphValueHistoriesService)
        {
            _lphValueHistoriesService = lphValueHistoriesService;
        }
    }
}
