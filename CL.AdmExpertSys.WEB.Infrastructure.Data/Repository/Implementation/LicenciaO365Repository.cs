using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;
using System.Data.Entity;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class LicenciaO365Repository : Repository<LICENCIAS_O365>, ILicenciaO365Repository
    {
        public LicenciaO365Repository(DbContext context)
            : base(context)
        {
        }
    }
}
