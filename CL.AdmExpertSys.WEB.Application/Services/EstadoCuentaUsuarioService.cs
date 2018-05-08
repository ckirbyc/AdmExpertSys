using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class EstadoCuentaUsuarioService : BaseService<ESTADO_CUENTA_USUARIO>, IEstadoCuentaUsuarioService
    {
        protected IEstadoCuentaUsuarioRepository Repo;

        public EstadoCuentaUsuarioService(IEstadoCuentaUsuarioRepository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
