using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.Models;
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
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoLicencia();
                var usuarioModificacion = SessionViewModel.Usuario.Nombre.Trim();
                var listaEstCuentaVmHilo = new List<object>
                {
                    listaEstUsr,
                    usuarioModificacion
                };

                _hiloEjecucion = new Thread(InciarProcesoHiloAsignarLicencia);
                _hiloEjecucion.Start(listaEstCuentaVmHilo);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = string.Empty,
                        Session = varSession
                    }
                };
            }
            catch (Exception ex)
            {
                HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(false, SessionViewModel.Usuario.Nombre.Trim());
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = false,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [SoapDocumentMethod(OneWay = true)]
        public void InciarProcesoHiloAsignarLicencia(object estadoCuentaHilo)
        {
            var usuarioModificacion = (string)estadoCuentaHilo.CastTo<List<object>>()[1];

            HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(true, usuarioModificacion);

            HomeSysWebFactory = new HomeSysWebFactory();            

            var estadoCuentaLista = (List<EstadoCuentaUsuarioVm>)estadoCuentaHilo.CastTo<List<object>>()[0];            

            foreach (var estUsr in estadoCuentaLista)
            {
                try
                {
                    var upnPrefijo = estUsr.CuentaAd.Trim() + estUsr.Dominio.Trim();
                    if (HomeSysWebFactory.ExisteUsuarioPortal(upnPrefijo))
                    {
                        var tieneLic = HomeSysWebFactory.ExisteLicenciaUsuarioPortal(upnPrefijo);
                        if (tieneLic)
                        {
                            estUsr.LicenciaAsignada = true;
                            HiloEstadoCuentaUsuario.ActualizarEstadoCuentaUsuario(estUsr);
                        }
                        else
                        {
                            if (HomeSysWebFactory.AsignarLicenciaUsuario(upnPrefijo, estUsr.LICENCIAS_O365.Codigo))
                            {
                                estUsr.LicenciaAsignada = true;
                                HiloEstadoCuentaUsuario.ActualizarEstadoCuentaUsuario(estUsr);
                            }
                        }                       
                    }
                }
                catch (Exception ex)
                {
                    HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(false, usuarioModificacion);
                    Utils.LogErrores(ex);
                }                
            }

            HiloEstadoAsignacionLicencia.ActualizarEstadoLicencia(false, usuarioModificacion);
        }
    }
}