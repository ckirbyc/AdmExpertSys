using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class UsuarioService : BaseService<USUARIO>, IUsuarioService
    {
        protected IUsuarioRepository Repo;

        public UsuarioService(IUsuarioRepository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
