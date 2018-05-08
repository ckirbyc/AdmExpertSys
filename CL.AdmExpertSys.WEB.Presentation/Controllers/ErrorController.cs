using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    /// <summary>
    /// Controlador de Error.
    /// </summary>
    [HandleError]
    public class ErrorController : Controller
    {
        /// <summary>
        /// Genera la página inicial de Error
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public ActionResult Index(string message)
        {
            ViewBag.Mensaje = message;
            return View();
        }

        public ActionResult IndexLogin(string mensajeError)
        {
            ViewBag.Mensaje = mensajeError;
            return View();
        }
    }
}
