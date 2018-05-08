
using CL.AdmExpertSys.WEB.Application.Contracts.Services;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class PerfilUsuarioFactory
    {
        protected IPerfilUsuarioService PerfilUsuarioService;

        public PerfilUsuarioFactory(
            IPerfilUsuarioService perfilUsuarioService)
        {
            PerfilUsuarioService = perfilUsuarioService;
        }
    }
}
