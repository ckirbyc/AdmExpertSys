using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Models;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class GruposController : BaseController
    {
        protected HomeSysWebFactory HomeSysWebFactory;
        protected Common CommonFactory;
        // GET: Grupos
        public ActionResult Index()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                CommonFactory = new Common();
                var ouAd = CommonFactory.GetAppSetting("RutaOuDominio");
                var listaGrupo = HomeSysWebFactory.ObtenerListadoGrupoAdByOuMantGroup(ouAd);
                return View(listaGrupo);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Grupos AD. Si el problema persiste contacte a soporte IT" });
            }            
        }

        public ActionResult Create()
        {
            HomeSysWebFactory = new HomeSysWebFactory();

            var objVm = HomeSysWebFactory.GetCreateViewGrupo();

            if (SessionViewModel.EstructuraArbolAd != null)
            {
                objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
            }
            else
            {
                var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                SessionViewModel.EstructuraArbolAd = estructura;
                objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
            }            

            return View(objVm);
        }

        [HttpPost]
        public ActionResult Create(GrupoAdVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                HomeSysWebFactory = new HomeSysWebFactory();
                if (!ModelState.IsValid)
                {                    
                    if (SessionViewModel.EstructuraArbolAd != null)
                    {
                        model.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                    }
                    else
                    {
                        var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                        SessionViewModel.EstructuraArbolAd = estructura;
                        model.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                    }
                    return View(model);
                }

                HomeSysWebFactory.CrearGrupoDistribucion(model);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar grabar registro de Grupo en el AD. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult Edit(string nombreGrupo)
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                var objVm = HomeSysWebFactory.GetGrupoDistribucion(nombreGrupo);

                if (SessionViewModel.EstructuraArbolAd != null)
                {
                    objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }
                else
                {
                    var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                    SessionViewModel.EstructuraArbolAd = estructura;
                    objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }

                return View(objVm);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar obtener registro de Grupo en el AD. Por favor reinténtelo más tarde." });
            }            
        }

        [HttpPost]
        public ActionResult Edit(GrupoAdVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                HomeSysWebFactory = new HomeSysWebFactory();
                if (!ModelState.IsValid)
                {
                    if (SessionViewModel.EstructuraArbolAd != null)
                    {
                        model.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                    }
                    else
                    {
                        var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                        SessionViewModel.EstructuraArbolAd = estructura;
                        model.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                    }                    
                    return View(model);
                }

                HomeSysWebFactory.ActualizarGrupoDistribucion(model);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar actualizar registro de Grupo en el AD. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult Delete(string nombreGrupo)
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                return View(HomeSysWebFactory.GetGrupoDistribucion(nombreGrupo));
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar eliminar registro de Grupo en el AD. Por favor reinténtelo más tarde." });
            }
        }

        [HttpPost]
        public ActionResult Delete(GrupoAdVm model)
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                HomeSysWebFactory.EliminarGrupoDistribucion(model);

                return RedirectToAction("Index");
            }            
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar eliminar Grupo del AD. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult VerPopupAsignarUsuarioGrupo(string nombreGrupo)
        {
            HomeSysWebFactory = new HomeSysWebFactory();
            var objVm = HomeSysWebFactory.GetGrupoDistribucion(nombreGrupo);

            if (SessionViewModel.EstructuraArbolAd != null)
            {
                objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
            }
            else
            {
                var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                SessionViewModel.EstructuraArbolAd = estructura;
                objVm.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
            }

            return View(objVm);
        }
    }
}