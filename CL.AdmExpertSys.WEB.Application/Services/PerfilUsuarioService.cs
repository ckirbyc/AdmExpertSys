using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class PerfilUsuarioService : BaseService<PERFIL_USUARIO>, IPerfilUsuarioService
    {
        protected IPerfilUsuarioRepository Repo;

        public PerfilUsuarioService(IPerfilUsuarioRepository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
