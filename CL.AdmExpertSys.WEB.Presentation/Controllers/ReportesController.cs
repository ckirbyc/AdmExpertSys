using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using System;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class ReportesController : BaseController
    {
        protected HomeSysWebFactory HomeSysWebFactory;
        // GET: Reportes
        public ActionResult Index()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();                               
                var listaCuenta = HomeSysWebFactory.ObtenerListaCuentaUsuario("N");
                ViewBag.processToken = Convert.ToInt64(DateTime.Now.Hour.ToString() + DateTime.Now.Minute + DateTime.Now.Second +
                                       DateTime.Now.Millisecond);
                return View(listaCuenta);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Reporte Cuenta Usuario AD. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult Licencia()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                var listaCuenta = HomeSysWebFactory.ObtenerListaCuentaUsuario("N");
                ViewBag.processToken = Convert.ToInt64(DateTime.Now.Hour.ToString() + DateTime.Now.Minute + DateTime.Now.Second +
                                       DateTime.Now.Millisecond);
                return View(listaCuenta);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Reporte Cuenta Usuario AD con Licencia. Si el problema persiste contacte a soporte IT" });
            }
        }

        public FileStreamResult ExportarExcel(string licencia, long processToken)
        {
            try
            {
                MakeProcessTokenCookie(processToken);

                HomeSysWebFactory = new HomeSysWebFactory();
                var libro = HomeSysWebFactory.ExportarArchivoExcelReporteCuentaUsuario(licencia);
                var memoryStream = new MemoryStream();
                libro.SaveAs(memoryStream);
                libro.Dispose();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                memoryStream.Flush();
                memoryStream.Position = 0;

                var nombreArchivo = string.Empty;
                if (licencia.Equals("N"))
                {
                    nombreArchivo = string.Format("ReporteCuentaUsuario.xlsx");
                }
                else
                {
                    nombreArchivo = string.Format("ReporteCuentaUsuarioLicencia.xlsx");
                }
                return File(memoryStream, "Reportes", nombreArchivo);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        private void MakeProcessTokenCookie(long processToken)
        {
            var cookie = new HttpCookie("processToken")
            {
                Value = processToken.ToString()
            };
            ControllerContext.HttpContext.Response.Cookies.Add(cookie);
        }
    }
}