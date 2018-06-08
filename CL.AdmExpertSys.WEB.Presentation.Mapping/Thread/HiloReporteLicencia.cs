using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Thread
{
    public static class HiloReporteLicencia
    {
        public static void CrearReporteLicenciaMasivo(List<UsuarioAd> objListaVm)
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        var listaRptLic = new List<REPORTE_LICENCIA>();

                        foreach (var usrAd in objListaVm)
                        {
                            var tip = new REPORTE_LICENCIA
                            {
                                Description = usrAd.Description,
                                DistinguishedName = usrAd.DistinguishedName,
                                InfoString = usrAd.InfoString,
                                Licenses = usrAd.Licenses,
                                Name = usrAd.Name,
                                SamAccountName = usrAd.SamAccountName,
                                UpnPrefijo = usrAd.UpnPrefijo
                            };
                            listaRptLic.Add(tip);
                        }

                        if (listaRptLic != null && listaRptLic.Count > 0)
                        {
                            entityContext.REPORTE_LICENCIA.AddRange(listaRptLic);
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

        public static bool ExistenRegistros()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    return entityContext.REPORTE_LICENCIA.Count() > 0 ? true : false;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void TruncateTablaReporteLicencia()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var objCtx = ((System.Data.Entity.Infrastructure.IObjectContextAdapter)entityContext).ObjectContext)
                    {
                        objCtx.ExecuteStoreCommand("TRUNCATE TABLE [REPORTE_LICENCIA]");                        
                    }
                }
            }
            catch (Exception)
            {                
                throw;
            }
        }

        public static List<UsuarioAd> GetListaReporteLicencia()
        {
            var listaUsrAd = new List<UsuarioAd>();
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    var listaObjBd = entityContext.REPORTE_LICENCIA.ToList();
                    foreach (var usrAdBd in listaObjBd)
                    {
                        var usrVm = new UsuarioAd
                        {
                            Description = usrAdBd.Description,
                            DistinguishedName = usrAdBd.DistinguishedName,
                            InfoString = usrAdBd.InfoString,
                            Licenses = usrAdBd.Licenses,
                            Name = usrAdBd.Name,
                            SamAccountName = usrAdBd.SamAccountName,
                            UpnPrefijo = usrAdBd.UpnPrefijo
                        };
                        listaUsrAd.Add(usrVm);
                    }

                    return listaUsrAd;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
