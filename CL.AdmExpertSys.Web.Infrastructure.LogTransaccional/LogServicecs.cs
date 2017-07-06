
using System;
using System.IO;

namespace CL.AdmExpertSys.Web.Infrastructure.LogTransaccional
{
    public static class LogServicecs
    {
        private static readonly object Locker = new Object();

        public static void EscribeLog(Exception ex)
        {
            try
            {
                lock (Locker)
                {
                    var sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\logfile.txt", true);
                    sw.WriteLine(DateTime.Now + ": " + ex.Source.Trim() + "; " + ex.Message.Trim());
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception)
            {
                EscribeLog("Error al insertar Exception.");
            }
        }

        public static void EscribeLog(string msj)
        {
            try
            {
                lock (Locker)
                {
                    var sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\logfile.txt", true);
                    sw.WriteLine(DateTime.Now + ": " + msj);
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }
        }
    }
}
