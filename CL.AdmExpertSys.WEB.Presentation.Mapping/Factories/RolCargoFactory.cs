using CL.AdmExpertSys.WEB.Application.Contracts.Services;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class RolCargoFactory
    {
        protected IRolCargoService RolCargoService;

        public RolCargoFactory(
            IRolCargoService rolCargoService)
        {
            RolCargoService = rolCargoService;
        }
    }
}
