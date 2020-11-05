using System;

namespace Fast.Domain.Interfaces.Repositories
{
	public interface IUnitOfWork : IDisposable
	{
		IRepositoryBase<TEntity> Repository<TEntity>() where TEntity : class;

		void SaveChanges();

		void RevertChanges();

	}
}
