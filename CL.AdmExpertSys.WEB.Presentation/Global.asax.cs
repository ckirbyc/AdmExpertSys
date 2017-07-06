using AutoMapper;
using CL.AdmExpertSys.WEB.Infrastructure.CompositionRoot;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Mapping;
using CL.AdmExpertSys.WEB.Presentation.Models;
using System;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace CL.AdmExpertSys.WEB.Presentation
{
    // Nota: para obtener instrucciones sobre cómo habilitar el modo clásico de IIS6 o IIS7, 
    // visite http://go.microsoft.com/?LinkId=9394801

    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            ContainerBootstrapper.RegisterTypes();
            Mapper.Initialize(cfg => cfg.AddProfile(new ViewModelProfile()));
            SqlServerTypes.Utilities.LoadNativeAssemblies(Server.MapPath("~/bin"));
            GlobalConfiguration.Configuration.Formatters.JsonFormatter.SerializerSettings.ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore;
        }

        protected void Session_Start()
        {
            try
            {
                var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                SessionViewModel.EstructuraArbolAd = estructura;
                //SessionViewModel.EstructuraArbolAd = null;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                //error al iniciar
                var urlHelper = new UrlHelper(Request.RequestContext);

                var url = urlHelper.Action("IndexLogin", "Error", new { message = "Error al iniciar la aplicación. Si el problema persiste contacte a soporte IT" });
                if (urlHelper.IsLocalUrl(url))
                {
                    Response.Redirect(url);
                }
                else
                {
                    url = urlHelper.Action("IndexLogin", "Error", new { message = "Error al iniciar la aplicación. Si el problema persiste contacte a soporte IT" });
                    Response.Redirect(url);
                }
            }
            HttpContext.Current.Session["UsuarioVM"] = "asd";
        }
    }
}