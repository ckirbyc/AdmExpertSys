using CL.AdmExpertSys.WEB.Application.Contracts.Services;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class LicenciaO365Factory
    {
        protected ILicenciaO365Service LicenciaO365Service;

        public LicenciaO365Factory(
            ILicenciaO365Service licenciaO365Service)
        {
            LicenciaO365Service = licenciaO365Service;
        }
    }
}
