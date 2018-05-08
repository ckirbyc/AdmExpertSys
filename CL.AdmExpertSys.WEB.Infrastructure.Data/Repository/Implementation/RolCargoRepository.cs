
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;
using System.Data.Entity;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class RolCargoRepository : Repository<ROL_CARGO>, IRolCargoRepository
    {
        public RolCargoRepository(DbContext context)
            : base(context)
        {
        }
    }
}
