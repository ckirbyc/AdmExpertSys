using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Thread
{
    public static class HiloEstadoCuentaUsuario
    {
        public static void ActualizarEstadoCuentaUsuario(EstadoCuentaUsuarioVm vm)
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        var objBd = entityContext.ESTADO_CUENTA_USUARIO.FirstOrDefault(x => x.IdEstadoCuentaUsuario == vm.IdEstadoCuentaUsuario);
                        if (objBd != null)
                        {
                            objBd.Apellidos = vm.Apellidos;
                            objBd.CodigoLicencia = vm.CodigoLicencia;
                            objBd.Correo = vm.Correo;
                            objBd.CreadoAd = vm.CreadoAd;
                            objBd.CuentaAd = vm.CuentaAd;
                            objBd.Descripcion = vm.Descripcion;
                            objBd.Dominio = vm.Dominio;
                            objBd.Eliminado = vm.Eliminado;
                            objBd.FechaBaja = vm.FechaBaja;
                            objBd.FechaCreacion = vm.FechaCreacion;
                            objBd.Habilitado = vm.Habilitado;
                            objBd.LicenciaAsignada = vm.LicenciaAsignada;
                            objBd.LicenciaId = vm.LicenciaId;
                            objBd.Nombres = vm.Nombres;
                            objBd.Sincronizado = vm.Sincronizado;
                            objBd.Vigente = vm.Vigente;

                            entityContext.SaveChanges();
                            dbContextTransaction.Commit();
                        }
                    }
                }                
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void CrearEstadoCuentaUsuario(EstadoCuentaUsuarioVm tipo)
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
                            LicenciaAsignada = false
                        };
                        entityContext.ESTADO_CUENTA_USUARIO.Add(tip);
                        entityContext.SaveChanges();
                        dbContextTransaction.Commit();                        
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }                      
        }
    }
}
