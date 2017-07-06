
using System;
using System.Linq;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using Pragma.Commons.Data.Patterns.Specification;
using Pragma.Commons.Domain;
using Pragma.Commons.Domain.Contracts;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class BaseService<T> : IBaseService<T> where T : Entity
    {
        private readonly IRepository<T> _repositorio;

        public BaseService(IRepository<T> repositorio)
        {
            _repositorio = repositorio;
        }

        public IQueryable<T> AllMatching(ISpecification<T> specification)
        {
            if (specification == null) throw new ArgumentNullException("specification");
            return _repositorio.AllMatching(specification);
        }

        public virtual void Create(T instance)
        {
            try
            {
                if (instance == null) throw new ArgumentNullException("instance");
                _repositorio.Create(instance);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }

        }

        public virtual void Save(T instance)
        {
            try
            {
                if (instance == null) throw new ArgumentNullException("instance");
                _repositorio.Create(instance);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }
        }

        public T FindById(long id)
        {
            return _repositorio.FindById(id);
        }

        public virtual T FindById(Func<T, bool> func)
        {
            return _repositorio.FindById(func);
        }

        public virtual IQueryable<T> FindAll()
        {
            return _repositorio.FindAll().Where(x => x.Vigente);
        }

        public virtual void Remove(T instance)
        {
            if (instance == null) throw new ArgumentNullException("instance");
            _repositorio.Remove(instance);
        }

        /// <summary>
        /// Obtiene todos los objetos, vigentes y no vigentes.
        /// </summary>
        /// <returns></returns>
        public virtual IQueryable<T> GetAll()
        {
            return _repositorio.FindAll();
        }

        public void Update(T instance)
        {
            try
            {
                if (instance == null) throw new ArgumentNullException("instance");
                _repositorio.Update(instance);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }
        }
    }
}
