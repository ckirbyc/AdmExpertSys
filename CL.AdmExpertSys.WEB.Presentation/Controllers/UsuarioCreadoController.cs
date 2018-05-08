using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using System;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class UsuarioCreadoController : BaseController
    {
        // GET: UsuarioCreado
        protected EstadoCuentaUsuarioFactory EstadoCuentaUsuarioFactory;        

        public UsuarioCreadoController()
        {
        }

        public UsuarioCreadoController(EstadoCuentaUsuarioFactory estadoCuentaUsuarioFactory)
        {            
            EstadoCuentaUsuarioFactory = estadoCuentaUsuarioFactory;            
        }

        public ActionResult Index()
        {
            try
            {                
                var model = EstadoCuentaUsuarioFactory.GetAllEstadoCuentaUsuario();
                return View(model);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página estado usuario. Si el problema persiste contacte a soporte IT" });
            }
        }
    }
}