using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Thread
{
    public static class HiloEstadoReporteLicencia
    {
        public static void ActualizarEstadoRptLicencia(bool valor)
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        var objBd = entityContext.ESTADO_REPORTE_LICENCIA.FirstOrDefault(x => x.Id == 1);
                        if (objBd != null)
                        {
                            objBd.Procesando = valor;

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

        public static bool EsProceso()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    return entityContext.ESTADO_REPORTE_LICENCIA.FirstOrDefault(x => x.Id == 1).Procesando;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
