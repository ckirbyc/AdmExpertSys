using CL.AdmExpertSys.WEB.Core.Domain.Model;
using System;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Thread
{
    public static class HiloEstadoSincronizacion
    {
        public static void ActualizarEstadoSync(bool valor, string userModifcacion, string modulo)
        {            
            using (var entityContext = new AdmSysWebEntities())
            {
                using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                {
                    var objBd = entityContext.ESTADO_SINCRONIZACION.FirstOrDefault(x => x.Id == 1);
                    if (objBd != null)
                    {
                        objBd.Sincronizando = valor;
                        objBd.FechaMvto = DateTime.Now;
                        objBd.UsuarioModificacion = userModifcacion;
                        objBd.Modulo = modulo;

                        entityContext.SaveChanges();
                        dbContextTransaction.Commit();
                    }
                }
            }          
        }

        public static bool EsSincronizacion()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    return entityContext.ESTADO_SINCRONIZACION.FirstOrDefault(x => x.Id == 1).Sincronizando;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
