using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class LicenciaO365Service : BaseService<LICENCIAS_O365>, ILicenciaO365Service
    {
        protected ILicenciaO365Repository Repo;

        public LicenciaO365Service(ILicenciaO365Repository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
