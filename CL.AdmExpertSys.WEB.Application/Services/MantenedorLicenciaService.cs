using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class MantenedorLicenciaService : BaseService<MANTENEDOR_LICENCIA>, IMantenedorLicenciaService
    {
        protected IMantenedorLicenciaRepository Repo;

        public MantenedorLicenciaService(IMantenedorLicenciaRepository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
