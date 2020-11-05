using Ninject;

namespace Fast.Infra.CrossCutting.IoC
{
	public class NinjectContainer
	{
		public StandardKernel GetServices()
		{
			return new StandardKernel(new NinjectRepositoryModule(),
				new NinjectServiceModule(),
				new NinjectAppServiceModule());
		}
	}
}
