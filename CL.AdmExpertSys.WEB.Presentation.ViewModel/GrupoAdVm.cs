
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Presentation.ViewModel.Validation;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class GrupoAdVm
    {
        [DisplayName(@"Identificador")]
        public int NumeroGrupo { get; set; }
        [DisplayName(@"Nombre")]
        [Required(ErrorMessage = @"Nombre es requerido")]
        public string NombreGrupo { get; set; }
        public string NombreGrupoAnterior { get; set; }
        [DisplayName(@"Ubicación")]
        [Required(ErrorMessage = @"Ubicación es requerido")]
        public string UbicacionGrupo { get; set; }
        [DisplayName(@"Correo")]
        [Required(ErrorMessage = @"Correo es requerido")]
        [ValidaCorreos]
        public string CorreoGrupo { get; set; }
        [DisplayName(@"Existe")]
        public bool ExisteGrupo { get; set; }
        [DisplayName(@"Tipo")]
        public string TipoGrupo { get; set; }
        [DisplayName(@"Descripción")]
        [Required(ErrorMessage = @"Descripción es requerido")]
        public string DescripcionGrupo { get; set; }
        public EstructuraArbolAd EstructuraArbolAd { get; set; }
        public bool Asociado { get; set; }        
        public List<UsuarioAd> ListaUsuarioAd { get; set; }
    }
}
