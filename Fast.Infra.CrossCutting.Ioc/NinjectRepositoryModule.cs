using Fast.Domain.Interfaces.Repositories;
using Fast.Infra.Data;
using Ninject.Modules;
using Ninject.Web.Common;

namespace Fast.Infra.CrossCutting.IoC
{
	public class NinjectRepositoryModule : NinjectModule
	{
		public override void Load()
		{
			Bind(typeof(IRepositoryBase<>)).To(typeof(RepositoryBase<>));
			Bind<IUnitOfWork>().To<UnitOfWork>().InRequestScope();
		}
	}
}
