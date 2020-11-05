using Fast.Domain.Interfaces.Repositories;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;

namespace Fast.Infra.Data
{
	public class RepositoryBase<TEntity> : IRepositoryBase<TEntity> where TEntity : class
	{
		protected readonly FastAppContext Db;
		protected IDbSet<TEntity> DbSet;

		public RepositoryBase(FastAppContext db)
		{
			Db = db;
			DbSet = Db.Set<TEntity>();
		}

		public void ExecuteSql(string query, object[] parameters)
		{
			Db.Database.ExecuteSqlCommand(query, parameters);
		}

		public void Add(TEntity obj)
		{
			DbSet.Add(obj);
		}

		public void AddRange(IEnumerable<TEntity> objList)
		{
			Db.Set<TEntity>().AddRange(objList);
		}

		public TEntity GetById(long id, bool asNoTracking = false)
		{
			if (asNoTracking)
			{
				var entity = DbSet.Find(id);
				if (entity == null)
				{
					return null;
				}
				else
				{
					Db.Entry(entity).State = EntityState.Detached;
					return entity;
				}
			}
			else
			{
				return DbSet.Find(id);
			}
		}

		public TEntity GetByName(string name, bool asNoTracking = false)
		{
			if (asNoTracking)
			{
				var entity = DbSet.Find(name);
				if (entity == null)
				{
					return null;
				}
				else
				{
					Db.Entry(entity).State = EntityState.Detached;
					return entity;
				}
			}
			else
			{
				return DbSet.Find(name);
			}
		}

		public ICollection<TEntity> GetAll()
		{
			return DbSet.ToList();
		}

		public void Update(TEntity obj)
		{
			Db.Entry(obj).State = EntityState.Modified;
		}

		public void Remove(TEntity obj)
		{
			DbSet.Remove(obj);
		}

		public void RemoveRange(IEnumerable<TEntity> objList)
		{
			Db.Set<TEntity>().RemoveRange(objList);
		}

		public TEntity Get(Expression<Func<TEntity, bool>> predicate)
		{
			return DbSet.Where(predicate).FirstOrDefault();
		}

		public TEntity Get(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false)
		{
			TEntity entity = DbSet.Where(predicate).FirstOrDefault();
			if (entity != null && asNoTracking)
				Db.Entry(entity).State = EntityState.Detached;

			return entity;
		}

		public TEntity GetLast(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false)
		{
			IEnumerable<TEntity> entities = DbSet.Where(predicate).ToList();
			TEntity entity = entities.LastOrDefault();

			if (entity != null && asNoTracking)
				Db.Entry(entity).State = EntityState.Detached;

			return entity;
		}

		public virtual TEntity Get(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes)
		{
			IQueryable<TEntity> query = DbSet;
			query = query.Where(predicate);
			query = includes.Aggregate(query, (current, include) => current.Include(include));
			TEntity result = query.FirstOrDefault();

			return result;
		}

		public IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate)
		{
			return DbSet.Where(predicate).ToList();
		}

		public IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes)
		{
			IQueryable<TEntity> query = DbSet;
			query = query.Where(predicate);
			query = includes.Aggregate(query, (current, include) => current.Include(include));

			IEnumerable<TEntity> result = query.ToList();
			return result;
		}

		public IEnumerable<TEntity> FindNoTracking(Expression<Func<TEntity, bool>> predicate)
		{
			var entity = DbSet.Where(predicate).ToList();
			foreach (var item in entity)
			{
				Db.Entry(item).State = EntityState.Detached;
			}
			return entity;
		}
	}
}
