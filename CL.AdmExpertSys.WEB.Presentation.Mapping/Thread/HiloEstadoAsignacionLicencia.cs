using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Thread
{
    public static class HiloEstadoAsignacionLicencia
    {
        public static void ActualizarEstadoLicencia(bool valor)
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        var objBd = entityContext.ESTADO_ASIGNACION_LICENCIA.FirstOrDefault(x => x.Id == 1);
                        if (objBd != null)
                        {
                            objBd.Asignando = valor;

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

        public static bool EsAsignacionLicencia()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    return entityContext.ESTADO_ASIGNACION_LICENCIA.FirstOrDefault(x => x.Id == 1).Asignando;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
