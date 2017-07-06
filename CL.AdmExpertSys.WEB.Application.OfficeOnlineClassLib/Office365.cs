
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace CL.AdmExpertSys.WEB.Application.OfficeOnlineClassLib
{
    /// <summary>
    /// Clase que contiene los métodos para realizar acciones en Office 365
    /// </summary>
    public class Office365
    {
        protected Common CommonServices;

        /// <summary>
        /// Clase para instanciar los parametros de Autorespuesta
        /// </summary>
        public class AutoReplyClass
        {
            public Boolean HasSetAutoReply;
            public DateTime StartTime;
            public DateTime EndTime;
            public string Message;
            public string User;
        }

        /// <summary>Indica si un usuario está licenciado en O365</summary>
        /// <param name="sUserName">El nombre de usuario a buscar</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>Retorna True en caso de que el usuario tiene licencia asignada</returns>
        public bool IsLicensedUser(string sUserName, out string sMess)
        {
            CommonServices = new Common();
            bool bReturn = false;
            sMess = string.Empty;
            try
            {
                // Create Initial Session State for runspace.
                InitialSessionState initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "MSOnline" });

                string sUser = CommonServices.GetAppSetting("usuarioO365");
                string sPwd = CommonServices.GetAppSetting("passwordO365");

                var securePass = new SecureString();
                foreach (char secureChar in sPwd)
                {
                    securePass.AppendChar(secureChar);
                }

                // Create credential object.
                var credential = new PSCredential(sUser, securePass);

                // Create command to connect office 365.
                var connectCommand = new Command("Connect-MsolService");
                connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));
                var getCommand = new Command("Get-MsolUser");
                getCommand.Parameters.Add((new CommandParameter("UserPrincipalName", sUserName)));
                using (Runspace psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
                {
                    // Open runspace.
                    psRunSpace.Open();
                    foreach (var com in new[] { connectCommand, getCommand })
                    {
                        var pipe = psRunSpace.CreatePipeline();
                        pipe.Commands.Add(com);
                        // Execute command and generate results and errors (if any).
                        Collection<PSObject> results = pipe.Invoke();
                        var error = pipe.Error.ReadToEnd();
                        Collection<PSObject> cOut;
                        if (error.Count > 0 && com == connectCommand)
                        {
                            sMess += " | Problema con el usuario y password del portal";
                            bReturn = false;
                        }
                        else cOut = results;
                        if (error.Count > 0 && com == getCommand)
                        {

                            sMess += " | Problema al obtener los datos";
                            bReturn = false;
                        }
                        else
                        {
                            cOut = results;
                            if (cOut != null && cOut.Count == 1)
                            {
                                foreach (PSObject user in cOut)
                                {
                                    if (user.Properties["isLicensed"] != null && user.Properties["isLicensed"].Value != null)
                                    {
                                        bReturn = user.Properties["isLicensed"].Value.ToString() == "True";
                                    }
                                    else bReturn = false;
                                }
                            }
                            else bReturn = false;
                        }
                    }
                    // Close the runspace.
                    psRunSpace.Close();
                    return bReturn;
                }
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        /// <summary>Indica si un usuario existe en O365</summary>
        /// <param name="sUserName">El nombre de usuario a buscar</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>Retorna True en caso de que el usuario existe en O365</returns>
        public bool UserExists(string sUserName, out string sMess)
        {
            CommonServices = new Common();
            bool bReturn = false;
            sMess = string.Empty;
            try
            {
                // Create Initial Session State for runspace.
                InitialSessionState initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "MSOnline" });

                string usuario = CommonServices.GetAppSetting("usuarioO365");
                string clave = CommonServices.GetAppSetting("passwordO365");
                var securePass = new SecureString();
                foreach (char secureChar in clave)
                {
                    securePass.AppendChar(secureChar);
                }

                // Create credential object.
                var credential = new PSCredential(usuario, securePass);

                // Create command to connect office 365.
                var connectCommand = new Command("Connect-MsolService");
                connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));
                var getCommand = new Command("Get-MsolUser");
                getCommand.Parameters.Add((new CommandParameter("UserPrincipalName", sUserName)));

                using (Runspace psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
                {
                    // Open runspace.
                    psRunSpace.Open();
                    foreach (var com in new[] { connectCommand, getCommand })
                    {
                        Collection<object> error;
                        using (var pipe = psRunSpace.CreatePipeline())
                        {
                            pipe.Commands.Add(com);
                            // Execute command and generate results and errors (if any).
                            //Collection<PSObject> results = pipe.Invoke();
                            pipe.Invoke();
                            error = pipe.Error.ReadToEnd();
                        }

                        if (error.Count > 0 && com == connectCommand)
                        {
                            sMess = @"Problema al conectarse al portal Office 365.";
                            bReturn = false;
                            break;
                        }
                        //else salida = results;
                        if (error.Count > 0 && com == getCommand)
                        {
                            sMess = @"Usuario consultado no existe en el portal Office 365. Espere que termine el proceso de sincronización o continue asignando grupos al usuario.";
                            bReturn = false;
                            break;
                        }
                        bReturn = true;
                    }
                    // Close the runspace.
                    psRunSpace.Close();
                    return bReturn;
                }
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        /// <summary>
        /// Asigna licencia del tipo sLicense al usuario indicado en sUPN
        /// </summary>
        /// <param name="sUpn">Usuario al que se le asigna la licencia indicada</param>
        /// <param name="sLicense">Tipo de licencia a asignar</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>
        /// Retorna True en caso de que se haya asignado la licencia al usuario indicado
        /// </returns>
        public bool AllocateLicense(string sUpn, string sLicense, out string sMess)
        {
            CommonServices = new Common();
            sMess = string.Empty;

            try
            {
                var initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "MSOnline" }); 
               
                string usuario = CommonServices.GetAppSetting("usuarioO365");
                string clave = CommonServices.GetAppSetting("passwordO365");
                var securePass = new SecureString();

                foreach (var secureChar in clave)
                {
                    securePass.AppendChar(secureChar);
                }

                // Create credential object.
                var credential = new PSCredential(usuario, securePass);

                // Create command to connect office 365.
                var connectCommand = new Command("Connect-MsolService");
                connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));

                var cmdLocation = new Command("Set-MsolUser");
                cmdLocation.Parameters.Add((new CommandParameter("UserPrincipalName", sUpn)));
                cmdLocation.Parameters.Add((new CommandParameter("UsageLocation", "CL")));

                var cmdlic = new Command("Set-MsolUserLicense");
                cmdlic.Parameters.Add((new CommandParameter("UserPrincipalName", sUpn)));
                cmdlic.Parameters.Add((new CommandParameter("AddLicenses", sLicense)));

                List<string> desLic;
                if (sLicense.Equals("agrosuper:ENTERPRISEPACK"))
                {
                    desLic = new List<string>
                    {
                        "PROJECTWORKMANAGEMENT", 
                        "INTUNE_O365", 
                        "YAMMER_ENTERPRISE", 
                        "RMS_S_ENTERPRISE", 
                        "SWAY", 
                        "OFFICESUBSCRIPTION", 
                        "MCOSTANDARD", 
                        "SHAREPOINTWAC", 
                        "SHAREPOINTENTERPRISE"
                    };
                }
                else if (sLicense.Equals("agrosuper:STANDARDPACK"))
                {
                    desLic = new List<string>
                    {
                        "SHAREPOINTWAC",
                        "PROJECTWORKMANAGEMENT",
                        "SWAY",
                        "INTUNE_O365",
                        "YAMMER_ENTERPRISE",
                        "MCOSTANDARD",
                        "SHAREPOINTSTANDARD"
                    };
                }
                else
                {
                    desLic = new List<string>
                    {
                        "INTUNE_O365"
                    };
                }


                var cmdLicDes = new Command("New-MsolLicenseOptions");
                cmdLicDes.Parameters.Add((new CommandParameter("AccountSkuId", sLicense)));
                cmdLicDes.Parameters.Add((new CommandParameter("DisabledPlans", desLic)));

                var lCommands = new List<Command> { connectCommand, cmdLocation, cmdlic, cmdLicDes };

                using (var psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
                {
                    // Open runspace.
                    psRunSpace.Open();

                    // Connect command
                    Collection<object> error1;
                    using (var pipe1 = psRunSpace.CreatePipeline())
                    {
                        pipe1.Commands.Add(lCommands[0]);
                        pipe1.Invoke();
                        error1 = pipe1.Error.ReadToEnd();
                    }
                    
                    if (error1.Count == 0)
                    {
                        Collection<object> error2;
                        using (var pipe2 = psRunSpace.CreatePipeline())
                        {
                            pipe2.Commands.Add(lCommands[1]);
                            pipe2.Invoke();
                            error2 = pipe2.Error.ReadToEnd();
                        }

                        if (error2.Any())
                        {
                            sMess += error2.ToString();
                            return false;
                        }

                        // Allocate License
                        Collection<PSObject> results3;
                        using (var pipe3 = psRunSpace.CreatePipeline())
                        {
                            pipe3.Commands.Add(lCommands[3]);
                            results3 = pipe3.Invoke();
                        }
                        PSObject userLicense = results3[0];
                        Collection<object> error3;
                        using (var pipe4 = psRunSpace.CreatePipeline())
                        {
                            var comandoLic = lCommands[2];
                            comandoLic.Parameters.Add((new CommandParameter("LicenseOptions", userLicense)));
                            pipe4.Commands.Add(comandoLic);
                            pipe4.Invoke();
                            error3 = pipe4.Error.ReadToEnd();
                        }
                        if (error3.Any())
                        {
                            sMess += error3.ToString();
                            return false;
                        }
                    }
                    else
                    {
                        sMess += error1.ToString();
                        return false;
                    }
                    psRunSpace.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        ///  <summary>Obtiene estructura para la autorespuesta de un usuario</summary>
        /// <param name="sUserName">El nombre de usuario a buscar</param>
        /// <returns>Retorna el objeto AutoReplyClass con la información de Auto Respuesta del Usuario inidcado</returns>
        public AutoReplyClass GetAutoReply(string sUserName)
        {
            CommonServices = new Common();
            AutoReplyClass arReturn = new AutoReplyClass();

            string usuario = CommonServices.GetAppSetting("usuarioO365");
            string clave = CommonServices.GetAppSetting("passwordO365");
            SecureString securePass = new SecureString();
            foreach (char secureChar in clave)
            {
                securePass.AppendChar(secureChar);
            }

            PSCredential credential = new PSCredential(usuario, securePass);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.SkipCACheck = true;
            connectionInfo.SkipCNCheck = true;

            connectionInfo.MaximumConnectionRedirectionCount = 4;
            Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
            try
            {
                runspace.Open();
                // Make a Get-Mailbox requst using the Server Argument
                Command gmGetMailbox = new Command("Get-MailboxAutoReplyConfiguration");
                gmGetMailbox.Parameters.Add("Identity", sUserName);
                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.Add(gmGetMailbox);
                Collection<PSObject> RsResultsresults = pipe.Invoke();
                //Dictionary<string, PSObject> gmResults = new Dictionary<string, PSObject>();
                if (RsResultsresults != null && RsResultsresults.Count == 1)
                {
                    foreach (PSObject user in RsResultsresults)
                    {
                        if (user.Properties["AutoReplyState"] != null && user.Properties["AutoReplyState"].Value != null)
                        {
                            Console.Write(user.Properties["AutoReplyState"].Value.ToString());
                            switch (user.Properties["AutoReplyState"].Value.ToString())
                            {
                                case "Disabled":
                                    arReturn.HasSetAutoReply = false;
                                    break;
                                case "Scheduled":
                                    arReturn.HasSetAutoReply = true;
                                    break;
                                case "Enabled":
                                    arReturn.HasSetAutoReply = true;
                                    break;
                                default:
                                    arReturn.HasSetAutoReply = false;
                                    break;
                            }
                            if (user.Properties["StartTime"] != null && user.Properties["StartTime"].Value != null) arReturn.StartTime = Convert.ToDateTime(user.Properties["StartTime"].Value.ToString());
                            if (user.Properties["EndTime"] != null && user.Properties["EndTime"].Value != null) arReturn.EndTime = Convert.ToDateTime(user.Properties["EndTime"].Value.ToString());
                            if (user.Properties["InternalMessage"] != null && user.Properties["InternalMessage"].Value != null) arReturn.Message = user.Properties["InternalMessage"].Value.ToString();
                        }
                    }
                }
                pipe.Stop();
                pipe.Dispose();
                runspace.Close();
                runspace.Dispose();
            }
            catch (Exception)
            {
                runspace.Close();
                runspace.Dispose();
                throw;
            }
            return arReturn;
        }

        /// <summary>Define la Auto Respuesta para el usuario inidcado</summary>
        /// <param name="usr">Objeto del tipo AutoreplyClass con la información de la Auto Respuesta</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>Retorna True en caso de efectuar la operación con éxito</returns>
        public bool SetAutorespuesta(AutoReplyClass usr, out string sMess)
        {
            CommonServices = new Common();
            sMess = string.Empty;
            string usuario = CommonServices.GetAppSetting("usuarioO365");
            string clave = CommonServices.GetAppSetting("passwordO365");
            SecureString securePass = new SecureString();
            foreach (char secureChar in clave)
            {
                securePass.AppendChar(secureChar);
            }

            PSCredential credential = new PSCredential(usuario, securePass);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.SkipCACheck = true;
            connectionInfo.SkipCNCheck = true;

            connectionInfo.MaximumConnectionRedirectionCount = 4;
            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            try
            {
                runspace.Open();
                // Make a Get-Mailbox requst using the Server Argument
                Command gmGetMailbox = new Command("Set-MailboxAutoReplyConfiguration");

                gmGetMailbox.Parameters.Add("Identity", usr.User);
                if (usr.HasSetAutoReply)
                {
                    gmGetMailbox.Parameters.Add("AutoReplyState", "Scheduled");
                    gmGetMailbox.Parameters.Add("StartTime", usr.StartTime);
                    gmGetMailbox.Parameters.Add("EndTime", usr.EndTime.AddDays(1));
                    gmGetMailbox.Parameters.Add("InternalMessage", usr.Message);
                    gmGetMailbox.Parameters.Add("ExternalMessage", usr.Message);
                }
                else gmGetMailbox.Parameters.Add("AutoReplyState", "Disabled");

                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.Add(gmGetMailbox);
                Collection<PSObject> RsResultsresults = pipe.Invoke();

                pipe.Stop();
                pipe.Dispose();
                runspace.Close();
                runspace.Dispose();

                if (RsResultsresults.Count > 0)
                {
                    sMess += " | Error";
                    return false;
                }
                return true;

            }
            catch (Exception)
            {
                runspace.Close();
                runspace.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Cambia el UPN del usuario iniacado en el O365
        /// </summary>
        /// <param name="oldUPN">UPN del usuario a modificar</param>
        /// <param name="newUPN"></param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>
        /// Retorna True en caso de efectuar la operación con éxito
        /// </returns>
        public bool ChangeUPN(string oldUPN, string newUPN, out string sMess)
        {
            CommonServices = new Common();
            sMess = string.Empty;
            bool bReturn = false;
            try
            {
                Collection<PSObject> cOut = null;
                // Create Initial Session State for runspace.
                InitialSessionState initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "MSOnline" });

                string usuario = CommonServices.GetAppSetting("usuarioO365");
                string clave = CommonServices.GetAppSetting("passwordO365");
                SecureString securePass = new SecureString();
                foreach (char secureChar in clave)
                {
                    securePass.AppendChar(secureChar);
                }

                // Create credential object.
                PSCredential credential = new PSCredential(usuario, securePass);
                // Create command to connect office 365.
                Command connectCommand = new Command("Connect-MsolService");
                connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));
                Command getCommand = new Command("Set-MsolUserPrincipalName");
                getCommand.Parameters.Add((new CommandParameter("UserPrincipalName", oldUPN)));
                getCommand.Parameters.Add((new CommandParameter("NewUserPrincipalName", newUPN)));
                using (Runspace psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
                {
                    // Open runspace.
                    psRunSpace.Open();
                    foreach (var com in new Command[] { connectCommand, getCommand })
                    {
                        var pipe = psRunSpace.CreatePipeline();
                        pipe.Commands.Add(com);
                        // Execute command and generate results and errors (if any).
                        Collection<PSObject> results = pipe.Invoke();
                        var error = pipe.Error.ReadToEnd();
                        if (error.Count > 0 && com == connectCommand)
                        {
                            sMess += " | Problem in login";
                            bReturn = false;
                        }
                        else cOut = results;
                        if (error.Count > 0 && com == getCommand)
                        {
                            sMess += "| Problem in setting data";
                            bReturn = false;
                        }
                        else
                        {
                            bReturn = true;
                        }
                    }
                    // Close the runspace.
                    psRunSpace.Close();
                    return bReturn;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Crea un buzón compartido
        /// </summary>
        /// <param name="sSharedMailBoxName">Nombre del buzón compartido</param>
        /// <param name="sSharedMailBoxAlias">Nombre del alias del buzón compartido</param>
        /// <param name="sPrimarySMTPAddr">Dirección SMTP primaria del buz{on
        /// compartido</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>
        /// Retorna True en caso de efectuar la operación con éxito
        /// </returns>
        public bool CreateSharedMailBox(string sSharedMailBoxName, string sSharedMailBoxAlias, string sPrimarySMTPAddr, out string sMess)
        {
            CommonServices = new Common();
            sMess = string.Empty;
            string usuario = CommonServices.GetAppSetting("usuarioO365");
            string clave = CommonServices.GetAppSetting("passwordO365");
            SecureString securePass = new SecureString();
            foreach (char secureChar in clave)
            {
                securePass.AppendChar(secureChar);
            }

            PSCredential credential = new PSCredential(usuario, securePass);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.SkipCACheck = true;
            connectionInfo.SkipCNCheck = true;

            connectionInfo.MaximumConnectionRedirectionCount = 4;
            Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
            try
            {
                runspace.Open();

                Command gmNewSharedMailbox = new Command("New-Mailbox");

                gmNewSharedMailbox.Parameters.Add("Name", sSharedMailBoxName);
                gmNewSharedMailbox.Parameters.Add("Alias", sSharedMailBoxAlias);
                gmNewSharedMailbox.Parameters.Add("PrimarySmtpAddress", sPrimarySMTPAddr);
                gmNewSharedMailbox.Parameters.Add("Shared", true);

                Command gmSetSharedMailbox = new Command("Set-Mailbox");

                gmSetSharedMailbox.Parameters.Add("Alias", sSharedMailBoxAlias);
                gmSetSharedMailbox.Parameters.Add("ProhibitSendReceiveQuota", "5GB");
                gmSetSharedMailbox.Parameters.Add("ProhibitSendQuota", "4.75GB");
                gmSetSharedMailbox.Parameters.Add("IssueWarningQuota", "4.5GB");


                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.Add(gmNewSharedMailbox);
                pipe.Commands.Add(gmSetSharedMailbox);
                Collection<PSObject> RsResultsresults = pipe.Invoke();

                pipe.Stop();
                pipe.Dispose();
                runspace.Close();
                runspace.Dispose();

                if (RsResultsresults.Count > 0)
                {
                    sMess += " | Error";
                    return false;
                }
                return true;

            }
            catch (Exception)
            {
                runspace.Close();
                runspace.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Chequea si un mailbox existe en Exchange Online
        /// </summary>
        /// <param name="sMailBoxIdentity">ID del Buzón</param>
        /// <param name="sPrimarySMTPAddr">Dirección SMTP primaria del buz{on
        /// compartido</param>
        /// <param name="sMess">Retorna condiciones de error</param>
        /// <returns>
        /// Retorna True en caso de existir el mailbox en Exchange Online
        /// </returns>
        public bool MailBoxExists(string sMailBoxIdentity, out string sMess)
        {
            CommonServices = new Common();
            sMess = string.Empty;
            string usuario = CommonServices.GetAppSetting("usuarioO365");
            string clave = CommonServices.GetAppSetting("passwordO365");
            SecureString securePass = new SecureString();
            foreach (char secureChar in clave)
            {
                securePass.AppendChar(secureChar);
            }

            PSCredential credential = new PSCredential(usuario, securePass);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.SkipCACheck = true;
            connectionInfo.SkipCNCheck = true;

            connectionInfo.MaximumConnectionRedirectionCount = 4;
            Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
            try
            {
                runspace.Open();

                Command gmGetMailbox = new Command("Get-Mailbox");

                gmGetMailbox.Parameters.Add("Identity", sMailBoxIdentity);



                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.Add(gmGetMailbox);
                Collection<PSObject> RsResultsresults = pipe.Invoke();

                pipe.Stop();
                pipe.Dispose();
                runspace.Close();
                runspace.Dispose();

                if (RsResultsresults.Count == 0)
                {
                    sMess += " | Sin resultados";
                    return false;
                }
                return true;

            }
            catch (Exception)
            {
                runspace.Close();
                runspace.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Sincronizar usuarios AD a O365
        /// </summary>
        /// <returns></returns>
        public bool ForzarDirSync()
        {
            //CommonServices = new Common();
            //var sRutaPsExec = CommonServices.GetAppSetting("RutaPsExec");
            //var sUserPsExec = CommonServices.GetAppSetting("UserPsExec");
            //var sPassPsExec = CommonServices.GetAppSetting("PassPsExec");
            //var sServerPsExec = CommonServices.GetAppSetting("ServerPsExec");
            //var sRutaBatPsExec = CommonServices.GetAppSetting("RutaBatPsExec");
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    //FileName = @"C:\\PSTools\\PsExec64.exe",
                    //Arguments = @"-d \\192.168.19.50 -u AS\mauricio.gonzalez -p inicio01 C:\\sync.bat"
                    FileName = @"C:\\sync.bat"
                };

                using (Process process = Process.Start(processInfo))
                {
                    if (process != null) process.WaitForExit();
                }
                //var processInfo = new ProcessStartInfo
                //{
                //    FileName = sRutaPsExec,
                //    Arguments = string.Format("-d \\{0} -u {1} -p {2} {3}", sServerPsExec, sUserPsExec, sPassPsExec, sRutaBatPsExec),
                //    UseShellExecute = false
                //};

                //using (var process = Process.Start(processInfo))
                //{
                //    if (process != null) process.WaitForExit();
                //}

                return true;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        /// <summary>Lista en duro de tipos de licencias y sus nombres de uso.</summary>
        /// <returns>Retorna un Hastable con los tipos de los tipos de licencias</returns>
        public Dictionary<string, string> Licenses()
        {
            var hOut = new Dictionary<string, string>
            {
              
                {"EXCHANGESTANDARD", "Exchange Online"},
                {"STANDARDPACK", "Plan E1"},
                {"ENTERPRISEPACK", "Plan E3"},
                {"POWER_BI_STANDARD", "Power BI"},
                {"EXCHANGEARCHIVE_ADDON","Archivado de Exchange Online" }
            };

            return hOut;
        }

        /// <summary>
        /// Metodo que obtiene la cantidad de licencias disponibles
        /// </summary>
        /// <returns></returns>
        public List<MsolAccountSku> ObtenerMsolAccountSku()
        {
            var listaMsolAccountSku = new List<MsolAccountSku>();
            CommonServices = new Common();
            try
            {
                Collection<PSObject> listMsolAccountSku = null;
                // Create Initial Session State for runspace.
                InitialSessionState initialSession = InitialSessionState.CreateDefault();
                initialSession.ImportPSModule(new[] { "MSOnline" });
                // Create credential object.
                string usuario = CommonServices.GetAppSetting("usuarioO365");
                string clave = CommonServices.GetAppSetting("passwordO365");
                var securePass = new SecureString();

                foreach (var secureChar in clave)
                {
                    securePass.AppendChar(secureChar);
                }

                // Create credential object.
                var credential = new PSCredential(usuario, securePass);
                // Create command to connect office 365.
                var connectCommand = new Command("Connect-MsolService");
                connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));
                // Create command to get office 365 users.
                var getLicenseCommand = new Command("Get-MsolAccountSku");
                getLicenseCommand.Parameters.Add((new CommandParameter("TenantId", "1259e2e7-e36c-4aa5-9256-ab0a2bd682bc")));
                //var lCommands = new List<Command> {connectCommand};

                using (Runspace psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
                {
                    // Open runspace.
                    psRunSpace.Open();
                    //Iterate through each command and executes it.
                    foreach (var com in new Command[] { connectCommand, getLicenseCommand })
                    {
                        using (var pipe = psRunSpace.CreatePipeline())
                        {
                            try
                            {
                                pipe.Commands.Add(com);
                                // Execute command and generate results and errors (if any).
                                Collection<PSObject> results = pipe.Invoke();
                                var error = pipe.Error.ReadToEnd();

                                if (error.Count > 0 && com == connectCommand)
                                {
                                    throw new Exception("Problema en el login : " + error[0]);
                                }
                                if (error.Count > 0 && com == getLicenseCommand)
                                {
                                    throw new Exception("Problema en el getting licenses : " + error[0]);
                                }

                                listMsolAccountSku = results;
                            }
                            catch (Exception ex) 
                            {
                                Utils.LogErrores(ex);
                            }
                        }
                    }
                    psRunSpace.Close();


                    if (listMsolAccountSku != null)
                    {
                        foreach (PSObject accountSku in listMsolAccountSku)
                        {
                            if (accountSku.Properties["SkuPartNumber"].Value.ToString().Equals("ENTERPRISEPACK") ||
                                accountSku.Properties["SkuPartNumber"].Value.ToString().Equals("STANDARDPACK") ||
                                accountSku.Properties["SkuPartNumber"].Value.ToString().Equals("EXCHANGESTANDARD"))
                            {
                                int disponibles = (Convert.ToInt32(accountSku.Properties["ActiveUnits"].Value) - Convert.ToInt32(accountSku.Properties["ConsumedUnits"].Value));
                                int total = (Convert.ToInt32(accountSku.Properties["ActiveUnits"].Value));
                                string plan = (accountSku.Properties["SkuPartNumber"].Value.ToString());

                                var sku = new MsolAccountSku
                                {
                                    CodigoLicencia = accountSku.Properties["SkuPartNumber"].Value.ToString(),
                                    MensajeCantidadLicencia = /*plan +*/ " " + disponibles + " de " + total
                                };
                                listaMsolAccountSku.Add(sku);
                            }
                        }
                    }
                }

                return listaMsolAccountSku;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw;
            }
        }
    }
}
