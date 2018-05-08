using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.Mvc;
using System.Web.Services.Protocols;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class LicCuentaController : BaseController
    {
        // GET: LicCuenta
        private static Thread _hiloEjecucion;

        protected EstadoCuentaUsuarioFactory EstadoCuentaUsuarioFactory;
        protected HomeSysWebFactory HomeSysWebFactory;

        public LicCuentaController()
        {
        }

        public LicCuentaController(EstadoCuentaUsuarioFactory estadoCuentaUsuarioFactory)
        {
            EstadoCuentaUsuarioFactory = estadoCuentaUsuarioFactory;
        }
        public ActionResult Index()
        {
            try
            {
                ViewBag.EstadoLicencia = HiloEstadoAsignacionLicencia.EsAsignacionLicencia();

                //Obtener todos los usuarios No poseen Licencia
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoLicencia();
                return View(listaEstUsr);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Asignar Licencias. Si el problema persiste contacte a soporte IT" });
            }
        }

        [HttpPost]
        public ActionResult AsignarLicencia()
        {
            try
            {
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoLicencia();
                var listaEstCuentaVmHilo = new List<object>
                {
                    listaEstUsr
                };

                _hiloEjecucion = new Thread(InciarProcesoHiloAsignarLicencia);
                _hiloEjecucion.Start(listaEstCuentaVmHilo);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = string.Empty
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
                        Validar = false,
                        Error = ex.Message
                    }
                };
            }
        }

        [SoapDocumentMethod(OneWay = true)]
        public void InciarProcesoHiloAsignarLicencia(object estadoCuentaHilo)
        {
            HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(true);

            HomeSysWebFactory = new HomeSysWebFactory();            

            var estadoCuentaLista = (List<EstadoCuentaUsuarioVm>)estadoCuentaHilo.CastTo<List<object>>()[0];            

            foreach (var estUsr in estadoCuentaLista)
            {
                if (HomeSysWebFactory.ExisteUsuarioPortal(estUsr.Correo))
                {
                    if (HomeSysWebFactory.AsignarLicenciaUsuario(estUsr.Correo, estUsr.LICENCIAS_O365.Codigo))
                    {
                        estUsr.LicenciaAsignada = true;
                        HiloEstadoCuentaUsuario.ActualizarEstadoCuentaUsuario(estUsr);
                    }                    
                }
            }

            HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(false);
        }
    }
}