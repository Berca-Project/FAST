using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Fast.Domain.Interfaces.Repositories
{
	public interface IRepositoryBase<TEntity> where TEntity : class
	{
		void Add(TEntity obj);
		void AddRange(IEnumerable<TEntity> objList);
		TEntity GetById(long id, bool asNoTracking = false);
		TEntity GetByName(string name, bool asNoTracking = false);
		ICollection<TEntity> GetAll();
		void Update(TEntity obj);
		void Remove(TEntity obj);
		void RemoveRange(IEnumerable<TEntity> objList);
		TEntity Get(Expression<Func<TEntity, bool>> predicate);
		TEntity Get(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false);
		TEntity GetLast(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false);
		TEntity Get(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes);
		IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate);
		IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes);
		IEnumerable<TEntity> FindNoTracking(Expression<Func<TEntity, bool>> predicate);
	}
}
