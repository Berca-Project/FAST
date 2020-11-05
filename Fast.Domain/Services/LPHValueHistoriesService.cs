using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Services
{
    public class LPHValueHistoriesService: ServiceBase<LPHValueHistories>, ILPHValueHistoriesService
    {
        public LPHValueHistoriesService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
