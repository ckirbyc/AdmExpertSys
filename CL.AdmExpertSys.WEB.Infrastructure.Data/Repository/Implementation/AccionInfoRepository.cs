
using System.Data.Entity;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class AccionInfoRepository : Repository<ACCION_INFO>, IAccionInfoRepository
    {
        public AccionInfoRepository(DbContext context)
            : base(context)
        {
        }
    }
}
