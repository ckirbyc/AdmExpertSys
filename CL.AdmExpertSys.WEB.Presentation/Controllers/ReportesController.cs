using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using System;
using System.IO;
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
                var listaCuenta = HomeSysWebFactory.ObtenerListaCuentaUsuario();
                return View(listaCuenta);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Reporte Cuenta Usuario AD. Si el problema persiste contacte a soporte IT" });
            }
        }

        public FileStreamResult ExportarExcel()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                var libro = HomeSysWebFactory.ExportarArchivoExcelReporteCuentaUsuario();
                var memoryStream = new MemoryStream();
                libro.SaveAs(memoryStream);
                libro.Dispose();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                memoryStream.Flush();
                memoryStream.Position = 0;

                var nombreArchivo = string.Format("ReporteCuentaUsuario.xlsx");
                return File(memoryStream, "Reportes", nombreArchivo);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }
    }
}