using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class PPReportYieldMCDietsService : ServiceBase<PPReportYieldMCDiets>, IPPReportYieldMCDietsService
    {
		public PPReportYieldMCDietsService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
    }
}
