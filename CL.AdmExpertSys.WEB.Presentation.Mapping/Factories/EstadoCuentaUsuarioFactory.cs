using AutoMapper;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using Pragma.Commons.Data.Patterns.Specification;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class EstadoCuentaUsuarioFactory
    {
        protected IEstadoCuentaUsuarioService EstadoCuentaUsuarioService;
        protected IMantenedorLicenciaService MantenedorLicenciaService;

        public EstadoCuentaUsuarioFactory(
            IEstadoCuentaUsuarioService estadoCuentaUsuarioService,
            IMantenedorLicenciaService mantenedorLicenciaService)
        {
            EstadoCuentaUsuarioService = estadoCuentaUsuarioService;
            MantenedorLicenciaService = mantenedorLicenciaService;
        }

        public EstadoCuentaUsuarioFactory() { }

        public IList<EstadoCuentaUsuarioVm> GetAllEstadoCuentaUsuario()
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.Habilitado && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<ESTADO_CUENTA_USUARIO>, IList<EstadoCuentaUsuarioVm>>(estadoBd);
            return vm.ToList();
        }

        public EstadoCuentaUsuarioVm GetEstadoCuentaUsuario(decimal id)
        {            
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.IdEstadoCuentaUsuario == id && x.Habilitado && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            if (estadoBd.Any()) {
                var estadoUno = estadoBd.FirstOrDefault();
                var retorno = Mapper.Map(estadoUno, new EstadoCuentaUsuarioVm());

                return retorno;
            }           

            return null;
        }

        public EstadoCuentaUsuarioVm GetEstadoCuentaUsuarioNoHabilitado(decimal id)
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.IdEstadoCuentaUsuario == id && x.Habilitado == false && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            if (estadoBd.Any())
            {
                var estadoUno = estadoBd.FirstOrDefault();
                var retorno = Mapper.Map(estadoUno, new EstadoCuentaUsuarioVm());

                return retorno;
            }

            return null;
        }

        public void CrearEstadoCuentaUsuario(EstadoCuentaUsuarioVm tipo)
        {
            var tip = new ESTADO_CUENTA_USUARIO
            {                
                Apellidos = tipo.Apellidos,
                CodigoLicencia = tipo.CodigoLicencia,
                Correo = tipo.Correo,
                CreadoAd = tipo.CreadoAd,
                CuentaAd = tipo.CuentaAd,
                Descripcion = tipo.Descripcion,
                Dominio = tipo.Dominio,
                Eliminado = tipo.Eliminado,
                FechaCreacion = tipo.FechaCreacion,
                Habilitado = tipo.Habilitado,
                LicenciaId = tipo.LicenciaId,
                Nombres = tipo.Nombres,
                Sincronizado = tipo.Sincronizado,
                Vigente = true,
                LicenciaAsignada = false,
                Clave = tipo.Clave.Trim()                                                                     
            };

            EstadoCuentaUsuarioService.Create(tip);
        }

        public void CrearEstadoCuentaUsuarioDirecto(EstadoCuentaUsuarioVm tipo)
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        var tip = new ESTADO_CUENTA_USUARIO
                        {
                            Apellidos = tipo.Apellidos,
                            CodigoLicencia = tipo.CodigoLicencia,
                            Correo = tipo.Correo,
                            CreadoAd = tipo.CreadoAd,
                            CuentaAd = tipo.CuentaAd,
                            Descripcion = tipo.Descripcion,
                            Dominio = tipo.Dominio,
                            Eliminado = tipo.Eliminado,
                            FechaCreacion = tipo.FechaCreacion,
                            Habilitado = tipo.Habilitado,
                            LicenciaId = tipo.LicenciaId,
                            Nombres = tipo.Nombres,
                            Sincronizado = tipo.Sincronizado,
                            Vigente = true,
                            LicenciaAsignada = tipo.LicenciaAsignada,
                            Clave = tipo.Clave.Trim()
                        };
                        entityContext.ESTADO_CUENTA_USUARIO.Add(tip);
                        entityContext.SaveChanges();
                        dbContextTransaction.Commit();
                    }
                }
            }
            catch (DbEntityValidationException dbEx)
            {
                Exception raise = dbEx;
                foreach (var validationErrors in dbEx.EntityValidationErrors)
                {
                    foreach (var validationError in validationErrors.ValidationErrors)
                    {
                        string message = string.Format("{0}:{1}",
                            validationErrors.Entry.Entity.ToString(),
                            validationError.ErrorMessage);
                        // raise a new exception nesting
                        // the current instance as InnerException
                        raise = new InvalidOperationException(message, raise);
                    }
                }
                throw raise;
            }
        }

        public void ActualizaEstadoCuentaUsuario(EstadoCuentaUsuarioVm tipo)
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.IdEstadoCuentaUsuario == tipo.IdEstadoCuentaUsuario);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var tip = estadoBd.FirstOrDefault();

            if (tip != null)
            {
                var estadoUsrBd = Mapper.Map(tipo, tip);
                EstadoCuentaUsuarioService.Update(estadoUsrBd);
            }            
        }

        public EstadoCuentaUsuarioVm GetObjetoEstadoCuentaUsuarioByCuenta(string cuentaAd)
        {
            try
            {
                var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CuentaAd == cuentaAd && x.Eliminado == false);
                var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

                if (estadoBd.Any())
                {
                    var retorno = Mapper.Map(estadoBd.FirstOrDefault(), new EstadoCuentaUsuarioVm());

                    return retorno;
                }

                return null;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }            
        }

        public EstadoCuentaUsuarioVm GetObjetoEstadoCuentaUsuarioAllByCuenta(string cuentaAd)
        {
            try
            {
                var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CuentaAd == cuentaAd && x.Eliminado == false);
                var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

                if (estadoBd.Any())
                {
                    var retorno = Mapper.Map(estadoBd.FirstOrDefault(), new EstadoCuentaUsuarioVm());

                    return retorno;
                }

                return null;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        public string GetCodigoLicenciaByUsuario(string cuentaAd) {
            try
            {
                var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CuentaAd == cuentaAd && x.Eliminado == false);
                var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

                if (estadoBd.Any()) {
                    var licId = estadoBd.FirstOrDefault().LicenciaId;
                    var codLic = string.Empty;
                    if (string.IsNullOrEmpty(estadoBd.FirstOrDefault().CodigoLicencia))
                    {
                        var especMantLic = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.LicenciaId == licId);
                        var estadoMantLicBd = MantenedorLicenciaService.AllMatching(especMantLic);
                        if (estadoMantLicBd.Any())
                        {
                            codLic = estadoMantLicBd.FirstOrDefault().Codigo.Trim();
                        }
                    }
                    else {
                        codLic = estadoBd.FirstOrDefault().CodigoLicencia;
                    }
                    
                    return codLic;
                }

                return string.Empty;
            }
            catch (Exception ex) {
                Utils.LogErrores(ex);
                return string.Empty;
            }
        }

        public string GetClaveCuentaByUsuario(string cuentaAd)
        {
            try
            {
                var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CuentaAd == cuentaAd && x.Eliminado == false);
                var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

                if (estadoBd.Any())
                {
                    return estadoBd.FirstOrDefault().Clave;
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return string.Empty;
            }
        }

        public IList<EstadoCuentaUsuarioVm> GetEstadoCuentaUsuarioNoSync()
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.Sincronizado == false && x.CreadoAd && x.LicenciaAsignada == false && x.Habilitado && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<ESTADO_CUENTA_USUARIO>, IList<EstadoCuentaUsuarioVm>>(estadoBd);
            return vm.ToList();
        }

        public IList<EstadoCuentaUsuarioVm> GetEstadoCuentaUsuarioNoLicencia()
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.LicenciaAsignada == false && x.CreadoAd && x.Sincronizado && x.Habilitado && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<ESTADO_CUENTA_USUARIO>, IList<EstadoCuentaUsuarioVm>>(estadoBd);
            return vm.ToList();
        }

        public IList<EstadoCuentaUsuarioVm> GetEstadoCuentaUsuarioDarBaja()
        {
            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CreadoAd && x.Habilitado == false && x.Eliminado == false);
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<ESTADO_CUENTA_USUARIO>, IList<EstadoCuentaUsuarioVm>>(estadoBd);
            return vm.ToList();
        }

        public IList<EstadoCuentaUsuarioVm> GetEstadoCuentaUsuarioDarBaja(string fechaDesde, string fechaHasta)
        {
            var fechaDesDate = Convert.ToDateTime(fechaDesde);
            var fechaHasDate = Convert.ToDateTime(fechaHasta);

            var espec = new DirectSpecification<ESTADO_CUENTA_USUARIO>(x => x.CreadoAd && x.Habilitado == false && x.Eliminado == false
            && (x.FechaBaja >= fechaDesDate) && (x.FechaBaja <= fechaHasDate));
            var estadoBd = EstadoCuentaUsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<ESTADO_CUENTA_USUARIO>, IList<EstadoCuentaUsuarioVm>>(estadoBd);
            return vm.ToList();
        }
    }
}
