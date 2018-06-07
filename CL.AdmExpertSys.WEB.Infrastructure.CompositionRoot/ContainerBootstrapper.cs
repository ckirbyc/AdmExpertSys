using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Application.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Microsoft.Practices.Unity;
using Microsoft.Practices.Unity.InterceptionExtension;
using System.Data.Entity;
using System.Web.Mvc;
using Unity.Mvc5;

namespace CL.AdmExpertSys.WEB.Infrastructure.CompositionRoot
{
    public class ContainerBootstrapper
    {
        public static void RegisterTypes()
        {
            IUnityContainer container = new UnityContainer();

            #region Connection
            //var connString = ConfigurationManager.ConnectionStrings["AdmSysWebEntities"].ConnectionString;
            //var conn = new SqlConnection {ConnectionString = connString};
            container.RegisterType<DbContext, AdmSysWebEntities>(new HierarchicalLifetimeManager(), new InjectionConstructor());
            #endregion

            #region ViewModelFactories            
            container.RegisterType<LogInfoFactory>();
            container.RegisterType<AccionInfoFactory>();
            container.RegisterType<EstadoCuentaUsuarioFactory>();
            container.RegisterType<LicenciaO365Factory>();
            container.RegisterType<MantenedorLicenciaFactory>();
            container.RegisterType<PerfilUsuarioFactory>();
            container.RegisterType<RolCargoFactory>();
            container.RegisterType<UsuarioFactory>();
            container.RegisterType<HomeSysWebFactory>();
            #endregion

            #region Repositories and services
            container.RegisterType<ILogInfoService, LogInfoService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IAccionInfoService, AccionInfoService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IPerfilUsuarioService, PerfilUsuarioService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IRolCargoService, RolCargoService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<ILicenciaO365Service, LicenciaO365Service>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IMantenedorLicenciaService, MantenedorLicenciaService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IEstadoCuentaUsuarioService, EstadoCuentaUsuarioService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<ITransversalService, TransversalService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IUsuarioService, UsuarioService>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());

            container.RegisterType<ILogInfoRepository, LogInfoRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IAccionInfoRepository, AccionInfoRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IPerfilUsuarioRepository, PerfilUsuarioRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IRolCargoRepository, RolCargoRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<ILicenciaO365Repository, LicenciaO365Repository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IMantenedorLicenciaRepository, MantenedorLicenciaRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IEstadoCuentaUsuarioRepository, EstadoCuentaUsuarioRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            container.RegisterType<IUsuarioRepository, UsuarioRepository>(new HierarchicalLifetimeManager(), new InterceptionBehavior<PolicyInjectionBehavior>());
            #endregion

            DependencyResolver.SetResolver(new UnityDependencyResolver(container));

            IConfigurationSource config = ConfigurationSourceFactory.Create();
            var logWriterFactory = new LogWriterFactory(config);

            var appLogger = logWriterFactory.Create();

            Logger.SetLogWriter(appLogger, false);

            var factory = new ExceptionPolicyFactory(config);
            var exManager = factory.CreateManager();
            ExceptionPolicy.SetExceptionManager(exManager, false);
        }
    }
}
