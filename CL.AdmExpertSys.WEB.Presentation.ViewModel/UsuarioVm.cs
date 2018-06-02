
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class UsuarioVm
    {
        [ScaffoldColumn(false)]
        public int Id { get; set; }

        [DisplayName(@"Descripción")]
        [StringLength(50)]
        [Required(ErrorMessage = @"El nombre es requerido")]
        public string Nombre { get; set; }

        public int Rut { get; set; }
        public string Dv { get; set; }
        public string Email { get; set; }
        public int? PerfilId { get; set; }
        public bool EsAdm { get; set; }
    }
}
