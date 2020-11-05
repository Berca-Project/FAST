using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Linq;
using System.Threading.Tasks;
using Fast.Domain.Interfaces.Repositories;

namespace Fast.Infra.Data
{
	public class UnitOfWork : IUnitOfWork, IDisposable
	{
		//private bool _disposed;
		//private readonly FastAppContext _context;
		public Dictionary<Type, object> Repositories = new Dictionary<Type, object>();
		private FastAppContext _context = new FastAppContext();

		public UnitOfWork()
		{
			//         
		}

		public IRepositoryBase<TEntity> Repository<TEntity>() where TEntity : class
		{
			if (Repositories.Keys.Contains(typeof(TEntity)))
			{
				return Repositories[typeof(TEntity)] as IRepositoryBase<TEntity>;
			}

			IRepositoryBase<TEntity> repository = new RepositoryBase<TEntity>(_context);

			Repositories.Add(typeof(TEntity), repository);

			return repository;
		}
		public void SaveChanges()
		{
			try
			{
				_context.SaveChanges();
			}
			catch (Exception e)
			{
				RevertChanges();
			}
		}

		public void RevertChanges()
		{
			//overwrite the existing context with a new, fresh one to revert all the changes
			_context = new FastAppContext();
		}

		private bool disposed = false;

		protected virtual void Dispose(bool disposing)
		{
			if (!this.disposed)
			{
				if (disposing)
				{
					_context.Dispose();
				}
			}
			this.disposed = true;
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

	}
}
