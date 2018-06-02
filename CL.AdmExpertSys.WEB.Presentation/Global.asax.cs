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
using CL.AdmExpertSys.WEB.Presentation.ViewModel;

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
                HttpContext.Current.Session["AdmExpertSys"] = "AdmExpertSys";                
                var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
                SessionViewModel.EstructuraArbolAd = estructura;                
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                //error al iniciar
                var urlHelper = new UrlHelper(Request.RequestContext);

                var url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = "Error al iniciar la aplicación. Si el problema persiste contacte a soporte IT" });
                if (urlHelper.IsLocalUrl(url))
                {
                    if (url != null) Response.Redirect(url);
                }
                else
                {
                    url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = "Error al iniciar la aplicación. Si el problema persiste contacte a soporte IT" });
                    if (url != null) Response.Redirect(url);
                }
            }            
        }

        protected void Application_Error(object sender, EventArgs e)
        {
            var contexto = Request.RequestContext;

            Exception exception = Server.GetLastError();

            Utils.LogErrores(exception);

            Response.Clear();

            var httpException = exception as HttpException;

            if (httpException != null)
            {
                var action = string.Empty;

                switch (httpException.GetHttpCode())
                {
                    case 404:
                        // page not found
                        action = "HttpError404";
                        break;
                    case 500:
                        // server error
                        action = "HttpError500";
                        break;
                    default:
                        action = "General";
                        break;
                }

                // clear error on server
                Server.ClearError();

                var urlHelper = new UrlHelper(Request.RequestContext);

                var url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = exception.Message });
                if (urlHelper.IsLocalUrl(url))
                {
                    Response.Redirect(url);
                }
                else
                {
                    url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = exception.Message });
                    Response.Redirect(url);
                }
            }
            else
            {
                // clear error on server
                Server.ClearError();
                string mensaje = string.Empty;

                if (!string.IsNullOrEmpty(exception.Message))
                {
                    mensaje += exception.Message;
                    mensaje += exception.InnerException != null ? "\n" + exception.InnerException.Message : string.Empty;
                }

                var urlHelper = new UrlHelper(Request.RequestContext);
                string url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = mensaje });

                if (urlHelper.IsLocalUrl(url))
                {
                    Response.Redirect(url);
                }
                else
                {
                    url = urlHelper.Action("IndexLogin", "Error", new { mensajeError = exception.Message });
                    Response.Redirect(url);
                }
            }
        }
    }
}