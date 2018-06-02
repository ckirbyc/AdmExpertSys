using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class UsuarioPerfilVm
    {
        [ScaffoldColumn(false)]
        [DisplayName(@"Identificador")]
        public decimal Id { get; set; }
        [DisplayName(@"Usuario")]
        [Required(ErrorMessage = @"Usuario es requerido")]
        public string Usuario1 { get; set; }
        [DisplayName(@"Perfil")]
        [Required(ErrorMessage = @"Perfil es requerido")]
        public int? PerfilId { get; set; }
        public IList<SelectListItem> PerfilSelectListItem { get; set; }
        public bool Visible { get; set; }
        [DisplayName(@"Administrador")]
        public bool EsAdm { get; set; }
        public virtual PERFIL_USUARIO PERFIL_USUARIO { get; set; }
    }
}
