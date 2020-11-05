using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Fast.Domain.Interfaces.Services
{
	public interface IServiceBase<TEntity> : IDisposable where TEntity : class
	{
		void Add(TEntity obj);
		void AddRange(IEnumerable<TEntity> objList);
		TEntity GetById(long id, bool asNoTracking = false);
		ICollection<TEntity> GetAll();
		void Update(TEntity obj);
		void Remove(TEntity obj);
		void RemoveRange(IEnumerable<TEntity> objList);
		TEntity Get(Expression<Func<TEntity, bool>> predicate);
		TEntity Get(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false);
		TEntity GetLast(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false);
		TEntity Get(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes);
		IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate);
		IEnumerable<TEntity> FindNoTracking(Expression<Func<TEntity, bool>> predicate);
		IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes);
		void Complete();
	}
}
