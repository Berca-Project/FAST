using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Fast.Domain.Services
{
	public class ServiceBase<TEntity> : IServiceBase<TEntity> where TEntity : class
	{
		private readonly IUnitOfWork _unitOfWork;
		private readonly IRepositoryBase<TEntity> _repository;

		public ServiceBase(IUnitOfWork unitOfWork)
		{
			_unitOfWork = unitOfWork;
			_repository = _unitOfWork.Repository<TEntity>();
		}

		public void Add(TEntity obj)
		{
			_repository.Add(obj);
		}

		public void AddRange(IEnumerable<TEntity> objList)
		{
			_repository.AddRange(objList);
		}

		public TEntity GetById(long id, bool asNoTracking = false)
		{
			return _repository.GetById(id, asNoTracking);
		}

		public ICollection<TEntity> GetAll()
		{
			return _repository.GetAll();
		}

		public void Update(TEntity obj)
		{
			_repository.Update(obj);
		}

		public void Remove(TEntity obj)
		{
			_repository.Remove(obj);
		}

		public void RemoveRange(IEnumerable<TEntity> objList)
		{
			_repository.RemoveRange(objList);
		}

		public TEntity Get(Expression<Func<TEntity, bool>> predicate)
		{
			return _repository.Get(predicate);
		}

		public TEntity Get(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false)
		{
			return _repository.Get(predicate, asNoTracking);
		}

		public TEntity GetLast(Expression<Func<TEntity, bool>> predicate, bool asNoTracking = false)
		{
			return _repository.GetLast(predicate, asNoTracking);
		}

		public TEntity Get(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes)
		{
			return _repository.Get(predicate, includes);
		}

		public IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate)
		{
			return _repository.Find(predicate);
		}

		public IEnumerable<TEntity> FindNoTracking(Expression<Func<TEntity, bool>> predicate)
		{
			return _repository.FindNoTracking(predicate);
		}

		public IEnumerable<TEntity> Find(Expression<Func<TEntity, bool>> predicate, ICollection<string> includes)
		{
			return _repository.Find(predicate, includes);
		}

		public void Complete()
		{
			//_unitOfWork.Complete();
			_unitOfWork.SaveChanges();
		}

		public void Dispose()
		{
			_unitOfWork.Dispose();
		}
	}
}
