using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class OuAdVm
    {
        [DisplayName(@"Nombre Ou")]
        public string Nombre { get; set; }

        [DisplayName(@"Ubicación")]
        public string Ldap { get; set; }

        [DisplayName(@"Atributo")]
        [Required(ErrorMessage = @"El Atributo es requerido")]
        public string Atributo { get; set; }
    }
}
