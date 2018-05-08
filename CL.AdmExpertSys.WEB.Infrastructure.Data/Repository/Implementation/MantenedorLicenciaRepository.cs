using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;
using System.Data.Entity;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class MantenedorLicenciaRepository : Repository<MANTENEDOR_LICENCIA>, IMantenedorLicenciaRepository
    {
        public MantenedorLicenciaRepository(DbContext context)
            : base(context)
        {
        }
    }
}
