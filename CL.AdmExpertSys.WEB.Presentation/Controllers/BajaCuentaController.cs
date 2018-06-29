using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class BajaCuentaController : BaseController
    {
        // GET: BajaCuenta

        protected EstadoCuentaUsuarioFactory EstadoCuentaUsuarioFactory;
        protected HomeSysWebFactory HomeSysWebFactory;

        public BajaCuentaController()
        {
        }

        public BajaCuentaController(EstadoCuentaUsuarioFactory estadoCuentaUsuarioFactory)
        {
            EstadoCuentaUsuarioFactory = estadoCuentaUsuarioFactory;
        }

        public ActionResult Index()
        {
            try
            {               
                //Obtener todos los usuarios deshabilitados para ser eliminados definitivamente del AD
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioDarBaja();
                return View(listaEstUsr);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Dar Baja Cuenta. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult GrillaIndex(List<EstadoCuentaUsuarioVm> model)
        {
            return PartialView("~/Views/BajaCuenta/_Index.cshtml", model);
        }

        public ActionResult FiltrarGrillaIndex(string fechaDesde, string fechaHasta)
        {
            IList<EstadoCuentaUsuarioVm> listaEstUsr = null;

            if (string.IsNullOrEmpty(fechaDesde) && string.IsNullOrEmpty(fechaHasta))
            {
                listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioDarBaja();
            }
            else
            {
                listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioDarBaja(fechaDesde, fechaHasta);
            }           

            var codHtml = RenderPartialViewToString("~/Views/BajaCuenta/_Index.cshtml", listaEstUsr);

            var response = new
            {
                Html = codHtml                
            };
            return Json(response, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult EliminarCuentaAd(string[] idCuentaAdArray)
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

                HomeSysWebFactory = new HomeSysWebFactory();
                var exito = false;
                var msgProceso = string.Empty;

                foreach (var idCuentaAdString in idCuentaAdArray)
                {
                    decimal idCuentaAd = Convert.ToDecimal(idCuentaAdString);
                    var estUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoHabilitado(idCuentaAd);

                    if (estUsr != null)
                    {
                        var usuarioAd = HomeSysWebFactory.ObtenerUsuarioExistente(estUsr.CuentaAd.Trim());
                        if (usuarioAd != null)
                        {
                            exito = HomeSysWebFactory.EliminarCuentaAd(estUsr.CuentaAd.Trim());                            
                        }
                        else {
                            exito = true;
                        }

                        if (exito == false)
                        {
                            msgProceso = @"La cuenta AD " + estUsr.CuentaAd + " no ha posido ser eliminado";
                            break;
                        }
                        else
                        {
                            estUsr.Eliminado = true;
                            estUsr.CreadoAd = false;
                            estUsr.Habilitado = false;
                            estUsr.LicenciaAsignada = false;
                            estUsr.Sincronizado = false;
                            estUsr.Vigente = false;
                            EstadoCuentaUsuarioFactory.ActualizaEstadoCuentaUsuario(estUsr);
                        }
                    }
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = exito,
                        Error = msgProceso,
                        Session = varSession
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
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }
    }
}