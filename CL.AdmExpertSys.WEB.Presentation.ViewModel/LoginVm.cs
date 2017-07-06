
namespace CL.AdmExpertSys.WEB.Presentation.ViewModel
{
    public class LoginVm
    {
        public string NombreUsuario { get; set; }
        public string ClaveUsuario { get; set; }
        public bool? EstaAutenticado { get; set; }
        public bool? TienePerfilesAsociados { get; set; }
        public bool UtilizarAutenticacion { get; set; }
    }
}
