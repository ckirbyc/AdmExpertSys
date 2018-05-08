using System.ComponentModel.DataAnnotations;
using System.Web;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class MantenedorLicenciaFileVm
    {
        [Required(ErrorMessage = @"Archivo es requerido")]
        public HttpPostedFileBase Archivo { get; set; }
    }
}
