
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class HomeSysWebVm
    {
        public string Nombres { get; set; }
        public string Apellidos { get; set; }
        public string NombreCompleto { get; set; }
        public int Rut { get; set; }
        public string Dv { get; set; }
        public string NombreUsuario { get; set; }
        public string Clave { get; set; }
        public string Correo { get; set; }
        public string Licencia { get; set; }
        public IList<SelectListItem> Licencias { get; set; }
        public string Ou { get; set; }
        public IList<SelectListItem> Ous { get; set; }
        public string UpnPrefijo { get; set; }
        public IList<SelectListItem> UpnPrefijoLista { get; set; }
        public string PatchOu { get; set; }
        public bool ExisteUsuario { get; set; }
        public List<MsolAccountSku> ListaAccountSkus { get; set; }
        public string NombreGrupo { get; set; }
        public string Descripcion { get; set; }
        [DisplayName(@"Código Licencia")]        
        public string CodigoLicencia { get; set; }
        public IList<SelectListItem> CodigoLicenciaLista { get; set; }
        public bool CambioPatchOu { get; set; }
        [DisplayName(@"Cuenta Genérica")]
        public bool Info { get; set; }
    }
}
