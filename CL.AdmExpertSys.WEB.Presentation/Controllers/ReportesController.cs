using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using System;
using System.IO;
using System.Threading;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Protocols;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class ReportesController : BaseController
    {
        private static Thread _hiloEjecucion;

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

                ViewBag.EstadoProceso = HiloEstadoReporteLicencia.EsProceso();
                ViewBag.ExistenRegistro = HiloReporteLicencia.ExistenRegistros();
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
                    HiloReporteLicencia.TruncateTablaReporteLicencia();
                }
                return File(memoryStream, "Reportes", nombreArchivo);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        [HttpPost]
        public ActionResult ProcesarDatos()
        {
            try
            {
                //Ejecuta Hilo para el proceso de sync de cuentas
                _hiloEjecucion = new Thread(InciarProcesoHiloReporteLicencia);
                _hiloEjecucion.Start();                

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true                        
                    }
                };
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = false
                    }
                };
            }
        }

        [SoapDocumentMethod(OneWay = true)]
        public void InciarProcesoHiloReporteLicencia()
        {
            try
            {
                HiloEstadoReporteLicencia.ActualizarEstadoRptLicencia(true);

                HomeSysWebFactory = new HomeSysWebFactory();
                var listaUsrAd = HomeSysWebFactory.ObtenerListaCuentaUsuarioLicense();
                if (listaUsrAd != null && listaUsrAd.Count > 0)
                {
                    HiloReporteLicencia.CrearReporteLicenciaMasivo(listaUsrAd);
                }

                HiloEstadoReporteLicencia.ActualizarEstadoRptLicencia(false);
            }
            catch (Exception ex)
            {
                HiloEstadoReporteLicencia.ActualizarEstadoRptLicencia(false);
                Utils.LogErrores(ex);
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