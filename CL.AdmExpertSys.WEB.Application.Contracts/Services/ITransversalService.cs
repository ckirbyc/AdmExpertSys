using System.Collections.Generic;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Application.Contracts.Services
{
    public interface ITransversalService
    {
        List<SelectListItem> GetSelectRolCarga();
        List<SelectListItem> GetSelectLicencia();
        List<SelectListItem> GetSelectPerfil();
    }
}
