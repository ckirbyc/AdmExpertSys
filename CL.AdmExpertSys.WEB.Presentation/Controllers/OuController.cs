using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Models;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class OuController : BaseController
    {
        protected HomeSysWebFactory HomeSysWebFactory;
        // GET: Ou
        public ActionResult Index()
        {
            try
            {   
                if (SessionViewModel.ListaOuAd == null)
                {
                    HomeSysWebFactory = new HomeSysWebFactory();
                    var listaOuAd = HomeSysWebFactory.GetListaOu();
                    SessionViewModel.ListaOuAd = listaOuAd;
                }               

                return View(SessionViewModel.ListaOuAd);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página de las OU. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult Edit(string nombreOu, string ldap)
        {
            try
            {
                var ouAd = new OuAdVm()
                {
                    Nombre = nombreOu,
                    Ldap = ldap
                };

                HomeSysWebFactory = new HomeSysWebFactory();
                ouAd.Atributo = HomeSysWebFactory.ObtenerAtributoOu(ldap);

                return View(ouAd);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página de edición OU. Si el problema persiste contacte a soporte IT" });
            }           
        }

        [HttpPost]
        public ActionResult Edit(OuAdVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                if (!ModelState.IsValid)
                {                    
                    return View(model);
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var exito = HomeSysWebFactory.ActualizarAttrOu(model.Ldap, model.Atributo.Trim());

                if(exito)
                    return RedirectToAction("Index");
                else
                    return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar guardar atributo de Unidad Organizativa. Por favor reinténtelo más tarde." });
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar guardar registro de Unidad Organizativa. Por favor reinténtelo más tarde." });
            }
        }
    }
}