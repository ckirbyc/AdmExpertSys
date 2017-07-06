
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling.PolicyInjection;
using Pragma.Commons.Data.Patterns.Specification;
using Pragma.Commons.Interception.Logging;
using System;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Application.Contracts.Services
{
    /// <summary>
    /// Interfaz Base que representa los metodos soportados por todos los repositorios
    /// </summary>
    /// <typeparam name="T">Entidad asociada al repositorio</typeparam>
    public interface IBaseService<T> where T : class
    {
        /// <summary>
        /// Get all elements of type TEntity that matching a
        /// Specification <paramref name="specification"/>
        /// </summary>
        /// <param name="specification">Specification that result meet</param>
        /// <returns></returns>
        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        IQueryable<T> AllMatching(ISpecification<T> specification);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        void Create(T instance);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        void Save(T instance);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        T FindById(long id);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        T FindById(Func<T, bool> func);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        IQueryable<T> FindAll();

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        IQueryable<T> GetAll();

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        void Remove(T instance);

        [Trace]
        [ExceptionCallHandler("BusinessLogicPolicy")]
        void Update(T instance);
    }
}
