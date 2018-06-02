using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class MantenedorLicenciaVm
    {
        [ScaffoldColumn(false)]
        [DisplayName(@"Identificador")]
        public decimal IdMantenedorLicencia { get; set; }
        [DisplayName(@"Código")]
        [Required(ErrorMessage = @"Código es requerido")]                
        public string Codigo { get; set; }
        [DisplayName(@"Rol Cargo")]
        [Required(ErrorMessage = @"Rol Cargo es requerido")]
        public decimal RolCargoId { get; set; }
        public IList<SelectListItem> RolCargoSelectListItem { get; set; }
        [DisplayName(@"Licencia")]
        [Required(ErrorMessage = @"Licencia es requerido")]
        public decimal LicenciaId { get; set; }
        public IList<SelectListItem> LicenciaSelectListItem { get; set; }
        public bool Vigente { get; set; }
        public LICENCIAS_O365 LICENCIAS_O365 { get; set; }
        public ROL_CARGO ROL_CARGO { get; set; }
    }
}
