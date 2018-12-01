using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CL.AdmExpertSys.WEB.Application.ADClassLib
{
    public class AdLib
    {
        #region Propiedades Privadas

        protected Common CommonServices;        
        private readonly string _sPublicDomain = string.Empty;
        private readonly string _sInternalDomain = string.Empty;
        private readonly string _sTenantDomain = string.Empty;
        private readonly string _sUserAdDomain = string.Empty;
        private readonly string _sPassAdDomain = string.Empty;
        private readonly string _sUpnPrefijo = string.Empty;
        private  readonly string _sLdapAs = string.Empty;
        private readonly string _sInternalDomainOu = string.Empty;
        private readonly string _sOuDeshabilitarUsr = string.Empty;
        private readonly string _sDominioAd = string.Empty;
        private readonly string _sRutaOuDominio = string.Empty;
        private readonly string _sNombreEmpresa = string.Empty;
        private readonly string _sLdapServer = string.Empty;
        private readonly string _sTenantDomainSmtp = string.Empty;
        private readonly string _sTenantDomainSmtpSecundario = string.Empty;
        private readonly string _sOuDeshabilitarComp = string.Empty;
        private readonly string _sRutaAllDominio = string.Empty;
        private readonly string _sIntentoCrearAd = string.Empty;

        #endregion

        public AdLib()
        {
            CommonServices = new Common();
            _sPublicDomain = CommonServices.GetAppSetting("PublicDomain");
            _sInternalDomain = CommonServices.GetAppSetting("InternalDomain");
            _sTenantDomain = CommonServices.GetAppSetting("TenantDomain");
            _sUserAdDomain = CommonServices.GetAppSetting("usuarioAd");
            _sPassAdDomain = CommonServices.GetAppSetting("passwordAd");
            _sUpnPrefijo = CommonServices.GetAppSetting("UpnPrefijo");
            _sLdapAs = CommonServices.GetAppSetting("LdapAs");
            _sInternalDomainOu = CommonServices.GetAppSetting("InternalDomainOu");
            _sOuDeshabilitarUsr = CommonServices.GetAppSetting("OuDeshabilitarUsr");
            _sDominioAd = CommonServices.GetAppSetting("dominioAd");
            _sRutaOuDominio = CommonServices.GetAppSetting("RutaOuDominio");
            _sNombreEmpresa = CommonServices.GetAppSetting("NombreEmpresa");
            _sLdapServer = CommonServices.GetAppSetting("LdapServidor");
            _sTenantDomainSmtp = CommonServices.GetAppSetting("TenantDomainSmtp");
            _sTenantDomainSmtpSecundario = CommonServices.GetAppSetting("TenantDomainSmtpSecundario");
            _sOuDeshabilitarComp = CommonServices.GetAppSetting("OuDeshabilitarComp");
            _sRutaAllDominio = CommonServices.GetAppSetting("RutaAllDominio");
            _sIntentoCrearAd = CommonServices.GetAppSetting("IntentoCrearAd");
        }

        #region Métodos de Validación


        /// <summary>
        /// Valida el nombre de usuario y contraseña para un usuario dado
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a validar</param>
        /// <param name="sPassword">La contraseña del usuario a validar</param>
        /// <returns>
        /// Retorna True si el usuario es válido
        /// </returns>
        public bool ValidateCredentials(string sUserName, string sPassword)
        {
            using (PrincipalContext oPrincipalContext = GetPrincipalContext())
            {
                return oPrincipalContext.ValidateCredentials(sUserName, sPassword, ContextOptions.Negotiate);
            }
        }


        /// <summary>
        ///  Valida si la cuenta de usuario está expirada
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a validar</param>
        /// <returns>
        /// Retorna true si ha expirado
        /// </returns>
        public bool IsUserExpired(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            if (oUserPrincipal.AccountExpirationDate != null)
            {
                return false;
            }
            return true;
        }


        /// <summary>
        /// Valida si el usuario existe en el AD
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a validar</param>
        /// <returns>
        /// Retorna true si el usuario existe
        /// </returns>
        public UsuarioAd IsUserExisting(string sUserName)
        {
            var oUserPrincipal = GetUser(sUserName);
            if (oUserPrincipal != null)
            {
                var info = ((DirectoryEntry)oUserPrincipal.GetUnderlyingObject()).Properties["info"];
                var infoString = string.Empty;
                var infoBool = false;
                if (info.Value != null)
                {
                    infoString = info.Value.ToString();
                    if (infoString.Equals("Cuenta Genérica"))
                    {
                        infoBool = true;
                    }
                }

                var usuarioAd = new UsuarioAd
                {
                    AccountExpirationDate = oUserPrincipal.AccountExpirationDate,
                    Description = oUserPrincipal.Description,
                    DisplayName = oUserPrincipal.DisplayName,
                    DistinguishedName = oUserPrincipal.DistinguishedName,
                    EmailAddress = oUserPrincipal.EmailAddress,
                    GivenName = oUserPrincipal.GivenName,
                    Guid = oUserPrincipal.Guid,
                    MiddleName = oUserPrincipal.MiddleName,
                    Name = oUserPrincipal.Name,
                    SamAccountName = oUserPrincipal.SamAccountName,
                    Surname = oUserPrincipal.Surname,
                    Enabled = oUserPrincipal.Enabled,
                    EstadoCuenta = oUserPrincipal.Enabled != null && oUserPrincipal.Enabled == true ? "Habilitado" : "No habilitado",
                    Info = infoBool                    
                };

                if (!string.IsNullOrEmpty(oUserPrincipal.UserPrincipalName))
                {
                    var upnPrefijo = oUserPrincipal.UserPrincipalName;
                    var startIndex = upnPrefijo.IndexOf("@");
                    var length = upnPrefijo.Length - startIndex;
                    usuarioAd.UpnPrefijo = upnPrefijo.Substring(startIndex, length);
                }

                return usuarioAd;
            }
            
            return null;
        }


        /// <summary>
        /// Valida si la cuenta de un usuario está bloqueada
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a validar</param>
        /// <returns>
        /// Retorna true si la cuenta está bloqueada
        /// </returns>
        public bool IsAccountLocked(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            return oUserPrincipal.IsAccountLockedOut();
        }
        #endregion

        #region Métodos de búsqueda



        /// <summary>
        /// Obtiene determinado usuario del AD
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a obtener</param>
        /// <returns>
        /// Retorna el objeto UserPrincipal
        /// </returns>
        public UserPrincipal GetUser(string sUserName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();
            UserPrincipal oUserPrincipal = UserPrincipal.FindByIdentity(oPrincipalContext, IdentityType.SamAccountName, sUserName);
            return oUserPrincipal;
        }

        public ComputerPrincipal GetComputerUser(string sComputerName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();
            ComputerPrincipal oComputerPrincipal = ComputerPrincipal.FindByIdentity(oPrincipalContext, IdentityType.Name, sComputerName);
            return oComputerPrincipal;                                 
        }

        /// <summary>
        /// Obtiene determinado grupo del AD
        /// </summary>
        /// <param name="sGroupName">El grupo a obtener</param>
        /// <returns>
        /// Retorna el objeto GroupPrincipal
        /// </returns>
        public GroupPrincipal GetGroup(string sGroupName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();
            var oGroupPrincipal = GroupPrincipal.FindByIdentity(oPrincipalContext, sGroupName);            
            return oGroupPrincipal;
        }
        /// <summary>
        /// Obtiene listado de todos los grupos pertenecientes a una OU especifica
        /// </summary>
        /// <param name="sOu">Ruta de la OU</param>
        /// <returns></returns>
        public List<GroupPrincipal> GetListGroupByOu(string sOuCompleto)
        {
            var listaGroup = new List<GroupPrincipal>();

            var sOu = sOuCompleto.Replace(_sLdapServer, string.Empty);

            PrincipalContext ctx = GetPrincipalContext(sOu);            
            using (GroupPrincipal objGroup = new GroupPrincipal(ctx))
            {
                objGroup.IsSecurityGroup = false;

                using (PrincipalSearcher pSearch = new PrincipalSearcher(objGroup))
                {
                    foreach (GroupPrincipal group in pSearch.FindAll())
                    {
                        var ubicGroup = group.DistinguishedName;
                        int startIndex = ubicGroup.IndexOf(",") + 1;
                        int length = ubicGroup.Length - startIndex;
                        var ubicClear = ubicGroup.Substring(startIndex, length);
                        if (sOu.Equals(ubicClear))
                        {
                            listaGroup.Add(group);
                        }
                    }                    
                    return listaGroup;
                }                       
            }                    
        }

        public PrincipalSearcher GetListGroupByOuMantGroup(string sOuCompleto)
        {           
            var sOu = sOuCompleto.Replace(_sLdapServer, string.Empty);
            PrincipalContext ctx = GetPrincipalContext(sOu);
            using (GroupPrincipal objGroup = new GroupPrincipal(ctx))
            {
                objGroup.IsSecurityGroup = false;
                PrincipalSearcher pSearch = new PrincipalSearcher(objGroup);                                    
                return pSearch;                
            }
        }

        #endregion

        #region Métodos de la cuenta de usuario


        /// <summary>
        /// Setea la contraseña de un usuario
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a setear</param>
        /// <param name="sNewPassword">La nueva contraseña a utilizar</param>
        /// <param name="sMessage">Mensaje de respuesta</param>
        public void SetUserPassword(string sUserName, string sNewPassword, out string sMessage)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                oUserPrincipal.SetPassword(sNewPassword);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }
            sMessage = null;
        }


        /// <summary>
        /// Habilita una cuenta de usuario deshabilitada
        /// </summary>
        /// <param name="sUserName">Nombre de usuario a habilitar</param>
        public void EnableUserAccount(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.Enabled = true;
            oUserPrincipal.Save();
        }


        /// <summary>
        /// Forzar deshabilitar una cuenta de usuario
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a deshabilitar</param>
        public bool DisableUserAccount(string sUserName)
        {
            try
            {
                using (UserPrincipal oUserPrincipal = GetUser(sUserName))
                {
                    oUserPrincipal.Enabled = false;
                    oUserPrincipal.Save();

                    string dnUser = oUserPrincipal.DistinguishedName;
                    var sLdapUser = _sLdapServer + dnUser;
                    var sLdapUserDest = _sLdapServer + _sOuDeshabilitarUsr;
                    using (var oLocation = new DirectoryEntry(sLdapUser, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                    {
                        using (var dLocation = new DirectoryEntry(sLdapUserDest, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                        {
                            oLocation.MoveTo(dLocation, oLocation.Name);
                            dLocation.Close();
                        }
                        oLocation.Close();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                var msgError = @"Error deshabilitar cuenta : " + sUserName + @", " + ex.Message;
                var exNew = new Exception(msgError);                
                Utils.LogErrores(exNew);
                return false;
            }
        }

        public bool DisableComputerAccount(string sComputerName)
        {
            try
            {
                using (ComputerPrincipal oComputerPrincipal = GetComputerUser(sComputerName))
                {
                    if (oComputerPrincipal != null)
                    {
                        oComputerPrincipal.Enabled = false;
                        oComputerPrincipal.Save();

                        string dnComputer = oComputerPrincipal.DistinguishedName;
                        var sLdapComputer = _sLdapServer + dnComputer;
                        var sLdapComputerDest = _sLdapServer + _sOuDeshabilitarComp;
                        using (var oLocation = new DirectoryEntry(sLdapComputer, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                        {
                            using (var dLocation = new DirectoryEntry(sLdapComputerDest, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                            {
                                oLocation.MoveTo(dLocation, oLocation.Name);
                                dLocation.Close();
                            }
                            oLocation.Close();
                        }
                    }                    
                    return true;
                }                
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return true;
            }
        }

        /// <summary>
        /// Forzar que expire la contraseña de un usuario
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a quien expirar la
        /// contraseña</param>
        public void ExpireUserPassword(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.ExpirePasswordNow();
            oUserPrincipal.Save();
        }


        /// <summary>
        /// Desbloquear una cuenta de usuario bloqueada
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a desbloquear</param>
        public void UnlockUserAccount(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.UnlockAccount();
            //oUserPrincipal.MSExchHideFromAddressLists=true;
            oUserPrincipal.Save();
        }


        /// <summary>
        /// Crear un nuevo usuario en el AD (*) Personalización
        /// </summary>
        /// <param name="ldapOu">La OU donde se desea crear el usuario</param>
        /// <param name="sUserName">El nombre de usuario</param>
        /// <param name="sPassword">La contraseña del nuevo usuario</param>
        /// <param name="sGivenName">El nombre para el nuevo usuario</param>
        /// <param name="sSurname">El apellido para el nuevo usuario</param>
        /// <param name="prefijoUpn"></param>
        /// <param name="passWord"></param>
        /// <param name="existeUsr"></param>
        /// <param name="descripcion"></param>
        /// <returns>
        /// Retorna el objeto UserPrincipal
        /// </returns>
        public UserPrincipal CreateNewUser(string ldapOu,
            string sUserName, string sPassword, string sGivenName, string sSurname, string prefijoUpn, string passWord, bool existeUsr, string descripcion, bool info)
        {
            string upn = sUserName + prefijoUpn;

            if (existeUsr == false)
            {
                var sOu = ldapOu.Replace(_sLdapServer, string.Empty);
                PrincipalContext oPrincipalContext = GetPrincipalContext(sOu);

                var oUserPrincipal = new UserPrincipal
                   (oPrincipalContext, sUserName, sPassword, true)
                { Enabled = true, PasswordNeverExpires = false };

                //Proxy Addresses                
                string emailOnmicrosoft = sUserName + "@" + _sTenantDomain;
                //string emailTransporte = "SMTP:"+ sUserName + "@agrosuper.mail.onmicrosoft.com";
                string emailTransporte = "SMTP:" + sUserName + '@' + _sTenantDomainSmtp;
                //string emailSecundario = sUserName + "@agrosuper.cl";
                string emailSecundario = sUserName + "@" + _sTenantDomainSmtpSecundario;

                string[] proxyaddresses = { "", "", "" };
                proxyaddresses[0] = "smtp:" + emailOnmicrosoft;
                proxyaddresses[1] = "SMTP:" + upn;
                proxyaddresses[2] = "smtp:" + emailSecundario;

                //User Log on Name
                oUserPrincipal.SamAccountName = sUserName;
                oUserPrincipal.UserPrincipalName = upn;
                oUserPrincipal.GivenName = sGivenName;
                oUserPrincipal.Surname = sSurname;
                oUserPrincipal.Name = sSurname + ", " + sGivenName;
                oUserPrincipal.MiddleName = sSurname;
                oUserPrincipal.DisplayName = sSurname + ", " + sGivenName;
                oUserPrincipal.EmailAddress = upn;
                oUserPrincipal.ExpirePasswordNow();
                oUserPrincipal.UnlockAccount();
                oUserPrincipal.Description = descripcion;
                oUserPrincipal.Save();

                using (DirectoryEntry ent = (DirectoryEntry)oUserPrincipal.GetUnderlyingObject())
                {
                    var infoString = string.Empty;
                    if (info)
                    {
                        infoString = @"Cuenta Genérica";
                    }
                    else
                    {
                        infoString = @"Cuenta Persona";
                    }                  

                    //Datos principales de la cuenta                    
                    ent.Invoke("SetPassword", passWord);
                    //Propiedades que son necesarias problar en AD para crear un usuario en Office 365
                    ent.Properties["proxyAddresses"].Add(proxyaddresses[0]);
                    ent.Properties["proxyAddresses"].Add(proxyaddresses[1]);
                    ent.Properties["proxyAddresses"].Add(proxyaddresses[2]);
                    ent.Properties["mailnickname"].Value = sUserName;
                    ent.Properties["targetAddress"].Value = emailTransporte;
                    ent.Properties["pwdLastSet"].Value = 0;
                    ent.Properties["info"].Value = infoString;

                    ent.CommitChanges();
                    ent.Close();
                };

                return oUserPrincipal;                
            }

            return GetUser(sUserName);
        }

        public bool UpdateUser(HomeSysWebVm usrData)
        {                        
            using (UserPrincipal oUserPrincipalAux = GetUser(usrData.NombreUsuario))
            {
                string dn = oUserPrincipalAux.DistinguishedName;
                var sLdapAsAux = _sLdapServer + dn;

                using (var eUserActual = new DirectoryEntry(sLdapAsAux, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                {
                    if (usrData.CambioPatchOu)
                    {
                        //Elimina los grupos asociados a la cuenta de usuario
                        var contMember = eUserActual.Properties["memberOf"].Count;
                        int equalsIndex, commaIndex;
                        for (int val = 0; val < contMember; val++)
                        {
                            var valMember = eUserActual.Properties["memberOf"][val].ToString();

                            equalsIndex = valMember.IndexOf("=", 1);
                            commaIndex = valMember.IndexOf(",", 1);

                            var groupName = valMember.Substring((equalsIndex + 1), (commaIndex - equalsIndex) - 1);

                            if (!string.IsNullOrEmpty(groupName))
                            {
                                RemoveUserFromGroup(usrData.NombreUsuario.Trim().ToLower(), groupName);
                            }
                        }
                    }                    

                    //Actualiza propiedades de la cuenta
                    CommonServices = new Common();
                    var sGivenName = CommonServices.UppercaseWords(usrData.Nombres.Trim().ToLower());
                    var sSurname = CommonServices.UppercaseWords(usrData.Apellidos.Trim().ToLower());
                    string upn = usrData.NombreUsuario.Trim() + usrData.UpnPrefijo.Trim();

                    if (usrData.Clave != "***")
                    {
                        eUserActual.Invoke("SetPassword", usrData.Clave.Trim());
                        //eUserActual.Properties["pwdLastSet"].Value = 0;
                    }                   

                    eUserActual.Properties["userAccountControl"].Value = 0x200;
                    eUserActual.Properties["userPrincipalName"].Value = upn;
                    eUserActual.Properties["givenName"].Value = sGivenName;
                    eUserActual.Properties["sn"].Value = sSurname;                    
                    eUserActual.Properties["middleName"].Value = sSurname;
                    eUserActual.Properties["displayName"].Value = sSurname + ", " + sGivenName;
                    eUserActual.Properties["mail"].Value = upn;
                    eUserActual.Properties["description"].Value = usrData.Descripcion.Trim();
                    eUserActual.Properties["mailnickname"].Value = usrData.NombreUsuario.Trim().ToLower();                                        

                    //Proxy Addresses                
                    string emailOnmicrosoft = usrData.NombreUsuario.Trim().ToLower() + "@" + _sTenantDomain;                    
                    string emailTransporte = "SMTP:" + usrData.NombreUsuario.Trim().ToLower() + '@' + _sTenantDomainSmtp;                    
                    string emailSecundario = usrData.NombreUsuario.Trim().ToLower() + "@" + _sTenantDomainSmtpSecundario;

                    string[] proxyaddresses = { "", "", "" };
                    proxyaddresses[0] = "smtp:" + emailOnmicrosoft;
                    proxyaddresses[1] = "SMTP:" + upn;
                    proxyaddresses[2] = "smtp:" + emailSecundario;

                    eUserActual.Properties["proxyAddresses"].Add(proxyaddresses[0]);
                    eUserActual.Properties["proxyAddresses"].Add(proxyaddresses[1]);
                    eUserActual.Properties["proxyAddresses"].Add(proxyaddresses[2]);

                    eUserActual.Properties["targetAddress"].Value = emailTransporte;

                    var infoString = string.Empty;
                    if (usrData.Info)
                    {
                        infoString = @"Cuenta Genérica";
                    }
                    else
                    {
                        infoString = @"Cuenta Persona";
                    }
                    eUserActual.Properties["info"].Value = infoString;

                    eUserActual.CommitChanges();
                    eUserActual.Close();                    

                    //Mueve cuenta de ubicacion
                    if (usrData.CambioPatchOu)
                    {
                        if (oUserPrincipalAux != null)
                        {                            
                            using (var eLocation = new DirectoryEntry(sLdapAsAux, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                            {
                                using (var nLocation = new DirectoryEntry(usrData.PatchOu, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                                {
                                    eLocation.MoveTo(nLocation);
                                    nLocation.Close();
                                }
                                eLocation.Close();
                            }
                            Task.Delay(TimeSpan.FromSeconds(15)).Wait();
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Setea una propiedad para un usuario determinado
        /// </summary>
        /// <param name="sUserName">Nombre de usuario para el que desea setear una propiedad
        /// de AD</param>
        /// <param name="sProperty">Propiedad de AD a setear (Ej: department)</param>
        /// <param name="sMessage">Mensaje de respuesta</param>
        public void SetUserProperty(string sUserName, string sProperty, /*string sPropertyValue,*/ out string sMessage)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                string dn = oUserPrincipal.DistinguishedName;

                DirectoryEntry ent = new DirectoryEntry("LDAP://" + _sInternalDomain + "/" + dn);
                //if (sPropertyValue != "" && sProperty != "") ent.Properties[sProperty].Value = sPropertyValue;
                ent.CommitChanges();
                sMessage = "";
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                sMessage = ex.Message;
            }

        }


        /// <summary>
        /// Borra un usuario del AD
        /// </summary>
        /// <param name="sUserName">Nombre de usuario que se desea borrar del AD</param>
        /// <returns>
        /// Retorna true si el borrado fue exitoso
        /// </returns>
        public bool DeleteUser(string sUserName)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);

                oUserPrincipal.Delete();
                return true;
            }
            catch(Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        #endregion

        #region Group Methods



        /// <summary>
        /// Crear un nuevo grupo en el AD
        /// </summary>
        /// <param name="sOu">La OU en que se desea guardar el nuevo grupo</param>
        /// <param name="sGroupName">El nombre del nuevo grupo</param>
        /// <param name="sDescription">Descripción para el nuevo grupo</param>
        /// <param name="oGroupScope">El Scope del nuevo grupo</param>
        /// <param name="bSecurityGroup">True si desea que este grupo sea un 'security
        /// group' o false si desea que sea 'distribution group'</param>
        /// <returns>
        /// Retorna el objeto GroupPrincipal
        /// </returns>
        public void CreateNewGroup(string sOu, string sGroupName, string sDescription, string sMail, GroupScope oGroupScope, bool bSecurityGroup)
        {
            var ldapOu = sOu.Replace(_sLdapServer, string.Empty);
            PrincipalContext oPrincipalContext = GetPrincipalContext(ldapOu);

            var oGroupPrincipal = new GroupPrincipal(oPrincipalContext, sGroupName)
            {
                Description = sDescription,
                GroupScope = oGroupScope,
                IsSecurityGroup = bSecurityGroup                
            };
            oGroupPrincipal.Save();

            oGroupPrincipal = GetGroup(sGroupName);
            ((DirectoryEntry)oGroupPrincipal.GetUnderlyingObject()).Properties["mail"].Value = sMail;

            oGroupPrincipal.Save();

            oGroupPrincipal.Dispose();
        }

        public void UpdateGroup(string sOu, string sGroupName, string sGroupNameAnt, string sDescription, string sMail, GroupScope oGroupScope, bool bSecurityGroup)
        {
            var ldapOu = sOu.Replace(_sLdapServer, string.Empty);            

            var oGroupPrincipal = GetGroup(sGroupNameAnt);
            var listaUsers = new List<string>();
            foreach (Principal objP in oGroupPrincipal.GetMembers(false))
            {
                listaUsers.Add(objP.SamAccountName);
            }

            oGroupPrincipal.Delete();
            oGroupPrincipal.Dispose();

            CreateNewGroup(sOu, sGroupName, sDescription, sMail, GroupScope.Global, false);

            foreach (var nameUsr in listaUsers)
            {
                AddUserToGroup(nameUsr, sGroupName);
            }            
        }

        public void DeleteGroup(string sGroupName)
        {
            var oGroupPrincipal = GetGroup(sGroupName);
            oGroupPrincipal.Delete();
            oGroupPrincipal.Dispose();
        }

        /// <summary>
        /// Agrega un usuario a un grupo dado
        /// </summary>
        /// <param name="sUserName">El nombre de usuario a agregar al grupo</param>
        /// <param name="sGroupName">El nombe del grupo donde agregar el usuario</param>
        /// <returns>
        /// Returns true if successful
        /// </returns>
        public bool AddUserToGroup(string sUserName, string sGroupName)
        {
            try
            {
                using (UserPrincipal oUserPrincipal = GetUser(sUserName))
                {
                    using (GroupPrincipal oGroupPrincipal = GetGroup(sGroupName))
                    {
                        if (oUserPrincipal != null && oGroupPrincipal != null)
                        {
                            if (!IsUserGroupMember(oUserPrincipal, oGroupPrincipal))
                            {
                                string userSid = string.Format("<SID={0}>", oUserPrincipal.Sid.ToString());
                                using (DirectoryEntry groupDirectoryEntry = (DirectoryEntry)oGroupPrincipal.GetUnderlyingObject())
                                {
                                    groupDirectoryEntry.Properties["member"].Add(userSid);
                                    groupDirectoryEntry.CommitChanges();
                                }                                                                                              
                                //oGroupPrincipal.Members.Add(oUserPrincipal);                                
                                //oGroupPrincipal.Save();
                            }
                        }
                    }
                }
                return true;
            }
            catch(Exception ex)
            {
                Utils.LogErrores(ex);
                throw;
            }
        }


        /// <summary>
        /// Elimina un usuario de un grupo
        /// </summary>
        /// <param name="sUserName">El nombre de usuario que se desea eliminar del
        /// grupo</param>
        /// <param name="sGroupName">El nombre del grupo de donde eliminar el
        /// usuario</param>
        /// <returns>
        /// Returns true if successful
        /// </returns>
        public bool RemoveUserFromGroup(string sUserName, string sGroupName)
        {
            try
            {
                using (UserPrincipal oUserPrincipal = GetUser(sUserName))
                {
                    using (GroupPrincipal oGroupPrincipal = GetGroup(sGroupName))
                    {
                        if (oUserPrincipal != null && oGroupPrincipal != null)
                        {
                            if (IsUserGroupMember(oUserPrincipal, oGroupPrincipal))
                            {
                                string userSid = string.Format("<SID={0}>", oUserPrincipal.Sid.ToString());
                                using (DirectoryEntry groupDirectoryEntry = (DirectoryEntry)oGroupPrincipal.GetUnderlyingObject())
                                {
                                    groupDirectoryEntry.Properties["member"].Remove(userSid);
                                    groupDirectoryEntry.CommitChanges();
                                }
                                //oGroupPrincipal.Members.Remove(oUserPrincipal);
                                //oGroupPrincipal.Save();
                            }
                        }
                    }
                }
                return true;
            }
            catch(Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }


        /// <summary>
        /// Verifica que un usuario pertenezca a un grupo
        /// </summary>
        /// <param name="oUserPrincipal">El nombre de usuario a validar</param>
        /// <param name="oGroupPrincipal">EL nombre del grupo para el cual se desae validad la
        /// pertenencia del usuario</param>
        /// <returns>
        /// Retorna true si el usuario pertenece al grupo
        /// </returns>
        public bool IsUserGroupMember(UserPrincipal oUserPrincipal, GroupPrincipal oGroupPrincipal)
        {
            if (oUserPrincipal != null && oGroupPrincipal != null)
            {
                List<string> lstUsuario = oGroupPrincipal.Members.Select(g => g.SamAccountName).ToList();
                var resultUsrExiste = lstUsuario.Where(x => x.ToString(CultureInfo.InvariantCulture).Equals(oUserPrincipal.SamAccountName)).ToList();
                if (resultUsrExiste.Any())
                {
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// Obtiene la lista de grupos a la que pertenece un usuario
        /// </summary>
        /// <param name="sUserName">Nombre de usuario para que se desea obtener los grupos a
        /// los que pertenece</param>
        /// <returns>
        /// Retorna un arraylist con los grupos a los que pertenece
        /// </returns>
        public ArrayList GetUserGroups(string sUserName)
        {
            var myItems = new ArrayList();
            UserPrincipal oUserPrincipal = GetUser(sUserName);

            PrincipalSearchResult<Principal> oPrincipalSearchResult = oUserPrincipal.GetGroups();

            foreach (Principal oResult in oPrincipalSearchResult)
            {
                myItems.Add(oResult.Name);
            }
            return myItems;
        }

        #endregion

        #region Helper Methods



        /// <summary>
        /// Obtiene el contexto principal
        /// </summary>
        /// <returns>
        /// Retorna el objeto PrincipalContext
        /// </returns>
        public PrincipalContext GetPrincipalContext()
        {
            var oPrincipalContext = new PrincipalContext(ContextType.Domain, _sInternalDomain, _sInternalDomainOu, _sUserAdDomain, _sPassAdDomain);
             return oPrincipalContext;
        }
        /// <summary>
        /// Gets the principal context on specified OU
        /// </summary>
        /// <param name="sOu">La OU a la que se desae obtener el contexto principal</param>
        /// <returns>
        /// Retorna el objeto PrincipalContext
        /// </returns>
        public PrincipalContext GetPrincipalContext(string sOu)
        {
            var oPrincipalContext = new PrincipalContext(ContextType.Domain, _sInternalDomain, sOu, _sUserAdDomain, _sPassAdDomain);
            return oPrincipalContext;
        }

        #endregion

        public ArrayList EnumerateOU(string OuDn)
        {
            ArrayList alObjects = new ArrayList();
            try
            {
                DirectoryEntry directoryObject = new DirectoryEntry("LDAP://" + OuDn);
                foreach (DirectoryEntry child in directoryObject.Children)
                {
                    string childPath = child.Path;
                    alObjects.Add(childPath.Remove(0, 7));
                    //remove the LDAP prefix from the path

                    child.Close();
                    child.Dispose();
                }
                directoryObject.Close();
                directoryObject.Dispose();
            }
            catch (DirectoryServicesCOMException e)
            {
                Console.WriteLine("An Error Occurred: " + e.Message);
            }
            return alObjects;
        }

        public Dictionary<string, string> OUs()
        {
            var hOut = new Dictionary<string, string>
            {
                {_sRutaOuDominio, _sNombreEmpresa}
            };
            return hOut;
        }

        public List<string> ObtenerPropiedadesUsuariosAd()
        {
            try
            {
                DirectoryEntry myLdapConnection = CrearEntradaDirectorio();
                var search = new DirectorySearcher(myLdapConnection);
                search.PropertiesToLoad.Add("cn");
                search.PropertiesToLoad.Add("mail");

                var listaPropiedades = new List<string>();
                SearchResultCollection allUsers = search.FindAll();
                foreach (SearchResult result in allUsers)
                {
                    if (result.Properties["cn"].Count > 0 && result.Properties["mail"].Count > 0)
                    {
                        listaPropiedades.Add(result.Properties["cn"][0].ToString());
                        listaPropiedades.Add(result.Properties["mail"][0].ToString());
                    }
                }

                return listaPropiedades;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }
            return null;
        }

        private DirectoryEntry CrearEntradaDirectorio()
        {
            // create and return new LDAP connection with desired settings 

            var ldapConnection = new DirectoryEntry
            {
                Path = _sLdapAs,
                AuthenticationType = AuthenticationTypes.Secure,
                Username = string.Format("{0}\\{1}",_sDominioAd,_sUserAdDomain),
                Password = _sPassAdDomain
                //Username = @"as\mauricio.gonzalez",
                //Password = @"inicio01"
            };
            return ldapConnection;
        }

        public List<Ou> EnumerateOu()
        {
            var listaOu = new List<Ou>();
            try
            {
                var directoryObject = CrearEntradaDirectorio();

                foreach (DirectoryEntry hijo in directoryObject.Children)
                {
                    if (hijo.Name.Contains("CN="))
                    {
                        continue;
                    }
                    var idOu1 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                    var objOu1 = new Ou { Nombre = hijo.Name, Nivel = 1, IdOu = idOu1, IdPadreOu = 0, Ldap = hijo.Path };
                    listaOu.Add(objOu1);
                    foreach (DirectoryEntry hijo2 in hijo.Children)
                    {
                        var idOu2 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                        var nombrePadre2 = hijo2.Parent.Name;
                        var existePadre2 = listaOu.Where(x => x.Nombre.Equals(nombrePadre2)).FirstOrDefault();
                        if (nombrePadre2.Contains("CN=") || existePadre2 == null || hijo2.Name.Contains("CN="))
                        {
                            continue;
                        }
                        var idPadre2 = listaOu.Where(x => x.Nombre.Equals(nombrePadre2)).FirstOrDefault().IdOu;
                        var objOu2 = new Ou { Nombre = hijo2.Name, Nivel = 2, IdOu = idOu2, IdPadreOu = idPadre2, Ldap = hijo2.Path };
                        listaOu.Add(objOu2);
                        foreach (DirectoryEntry hijo3 in hijo2.Children)
                        {
                            var idOu3 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                            var nombrePadre3 = hijo3.Parent.Name;
                            var existePadre3 = listaOu.Where(x => x.Nombre.Equals(nombrePadre3)).FirstOrDefault();
                            if (nombrePadre3.Contains("CN=") || existePadre3 == null || hijo3.Name.Contains("CN="))
                            {
                                continue;
                            }
                            var idPadre3 = listaOu.Where(x => x.Nombre.Equals(nombrePadre3)).FirstOrDefault().IdOu;
                            var objOu3 = new Ou { Nombre = hijo3.Name, Nivel = 3, IdOu = idOu3, IdPadreOu = idPadre3, Ldap = hijo3.Path };
                            listaOu.Add(objOu3);
                            foreach (DirectoryEntry hijo4 in hijo3.Children)
                            {
                                var idOu4 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                                var nombrePadre4 = hijo4.Parent.Name;
                                var existePadre4 = listaOu.Where(x => x.Nombre.Equals(nombrePadre4)).FirstOrDefault();
                                if (nombrePadre4.Contains("CN=") || existePadre4 == null || hijo4.Name.Contains("CN="))
                                {
                                    continue;
                                }
                                var idPadre4 = listaOu.Where(x => x.Nombre.Equals(nombrePadre4)).FirstOrDefault().IdOu;
                                var objOu4 = new Ou { Nombre = hijo4.Name, Nivel = 4, IdOu = idOu4, IdPadreOu = idPadre4, Ldap = hijo4.Path };
                                listaOu.Add(objOu4);
                                foreach (DirectoryEntry hijo5 in hijo4.Children)
                                {
                                    var idOu5 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                                    var nombrePadre5 = hijo5.Parent.Name;
                                    var existePadre5 = listaOu.Where(x => x.Nombre.Equals(nombrePadre5)).FirstOrDefault();
                                    if (nombrePadre5.Contains("CN=") || existePadre5 == null || hijo5.Name.Contains("CN="))
                                    {
                                        continue;
                                    }
                                    var idPadre5 = listaOu.Where(x => x.Nombre.Equals(nombrePadre5)).FirstOrDefault().IdOu;
                                    var objOu5 = new Ou { Nombre = hijo5.Name, Nivel = 5, IdOu = idOu5, IdPadreOu = idPadre5, Ldap = hijo5.Path };
                                    listaOu.Add(objOu5);
                                    foreach (DirectoryEntry hijo6 in hijo5.Children)
                                    {
                                        var idOu6 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                                        var nombrePadre6 = hijo6.Parent.Name;
                                        var existePadre6 = listaOu.Where(x => x.Nombre.Equals(nombrePadre6)).FirstOrDefault();
                                        if (nombrePadre6.Contains("CN=") || existePadre6 == null || hijo6.Name.Contains("CN="))
                                        {
                                            continue;
                                        }
                                        var idPadre6 = listaOu.Where(x => x.Nombre.Equals(nombrePadre6)).FirstOrDefault().IdOu;
                                        var objOu6 = new Ou { Nombre = hijo6.Name, Nivel = 6, IdOu = idOu6, IdPadreOu = idPadre6, Ldap = hijo6.Path };
                                        listaOu.Add(objOu6);
                                        foreach (DirectoryEntry hijo7 in hijo6.Children)
                                        {
                                            var idOu7 = listaOu.Count == 0 ? 1 : listaOu.Max(x => x.IdOu) + 1;
                                            var nombrePadre7 = hijo7.Parent.Name;
                                            var existePadre7 = listaOu.Where(x => x.Nombre.Equals(nombrePadre7)).FirstOrDefault();
                                            if (nombrePadre7.Contains("CN=") || existePadre7 == null || hijo7.Name.Contains("CN="))
                                            {
                                                continue;
                                            }
                                            var idPadre7 = listaOu.Where(x => x.Nombre.Equals(nombrePadre7)).FirstOrDefault().IdOu;
                                            var objOu7 = new Ou { Nombre = hijo7.Name, Nivel = 7, IdOu = idOu7, IdPadreOu = idPadre7, Ldap = hijo7.Path };
                                            listaOu.Add(objOu7);
                                            hijo7.Close();
                                            hijo7.Dispose();
                                        }
                                        hijo6.Close();
                                        hijo6.Dispose();
                                    }
                                    hijo5.Close();
                                    hijo5.Dispose();
                                }
                                hijo4.Close();
                                hijo4.Dispose();
                            }
                            hijo3.Close();
                            hijo3.Dispose();
                        }
                        hijo2.Close();
                        hijo2.Dispose();
                    }

                    hijo.Close();
                    hijo.Dispose();
                }
                directoryObject.Close();
                directoryObject.Dispose();
            }
            catch (DirectoryServicesCOMException ex)
            {
                Utils.LogErrores(ex);
            }
            return listaOu;
        }

        public Dictionary<string, string> GetUpnPrefijo()
        {
            return _sUpnPrefijo.Split(';').ToDictionary(upnPref => upnPref);
        }

        public List<UsuarioAd> GetListAccountUsers(string generarInfo, string generarAttr)
        {
            var listaUser = new List<UsuarioAd>();
            using (PrincipalContext oPrincipalContext = GetPrincipalContext(_sRutaOuDominio))
            {
                using (UserPrincipal objUser = new UserPrincipal(oPrincipalContext))
                {
                    objUser.Enabled = true;                    
                    using (PrincipalSearcher pSearch = new PrincipalSearcher(objUser))
                    {
                        foreach (UserPrincipal oUserPrincipal in pSearch.FindAll())
                        {
                            if (oUserPrincipal != null && !string.IsNullOrEmpty(oUserPrincipal.UserPrincipalName))
                            {
                                var upnPrefijoFinal = string.Empty;
                                if (oUserPrincipal.EmailAddress != null)
                                {
                                    var upnPrefijo = oUserPrincipal.EmailAddress;
                                    var startIndex = upnPrefijo.IndexOf("@");
                                    var length = upnPrefijo.Length - startIndex;
                                    upnPrefijoFinal = upnPrefijo.Substring(startIndex, length);
                                }
                                
                                var infoString = string.Empty;
                                if (generarInfo.Equals("S"))
                                {
                                    var info = ((DirectoryEntry)oUserPrincipal.GetUnderlyingObject()).Properties["info"];
                                    if (info.Value != null)
                                    {
                                        infoString = info.Value.ToString();
                                    }
                                }

                                var attrOu = string.Empty;
                                if (generarAttr.Equals("S"))
                                {
                                    var rutaPadreOu = ((DirectoryEntry)oUserPrincipal.GetUnderlyingObject()).Parent.Path;
                                    using (var eRutaPadre = new DirectoryEntry(rutaPadreOu, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
                                    {
                                        if (eRutaPadre.Properties["businessCategory"].Value != null)
                                        {
                                            attrOu = Convert.ToString(eRutaPadre.Properties["businessCategory"].Value);
                                        }
                                    };
                                }                                

                                var usuarioAd = new UsuarioAd
                                {
                                    AccountExpirationDate = oUserPrincipal.AccountExpirationDate,
                                    Description = oUserPrincipal.Description,
                                    DisplayName = oUserPrincipal.DisplayName,
                                    DistinguishedName = oUserPrincipal.DistinguishedName,
                                    EmailAddress = oUserPrincipal.EmailAddress,
                                    GivenName = oUserPrincipal.GivenName,
                                    Guid = oUserPrincipal.Guid,
                                    MiddleName = oUserPrincipal.MiddleName,
                                    Name = oUserPrincipal.Name,
                                    SamAccountName = oUserPrincipal.SamAccountName,
                                    Surname = oUserPrincipal.Surname,
                                    Enabled = oUserPrincipal.Enabled,
                                    EstadoCuenta = oUserPrincipal.Enabled != null && oUserPrincipal.Enabled == true ? "Habilitado" : "No habilitado",                                    
                                    UpnPrefijo = upnPrefijoFinal,
                                    InfoString = infoString,
                                    AttrOu = attrOu
                                };
                                listaUser.Add(usuarioAd);
                            }
                        }

                        return listaUser;
                    }                                      
                }                             
            }                                
        }

        public List<UsuarioAd> GetListAccountAllUsersAd(string generarInfo)
        {
            var listaUser = new List<UsuarioAd>();
            using (PrincipalContext oPrincipalContext = GetPrincipalContext(_sRutaAllDominio))
            {
                using (UserPrincipal objUser = new UserPrincipal(oPrincipalContext))
                {
                    objUser.Enabled = true;                    
                    using (PrincipalSearcher pSearch = new PrincipalSearcher(objUser))
                    {
                        foreach (UserPrincipal oUserPrincipal in pSearch.FindAll())
                        {
                            if (oUserPrincipal != null && !string.IsNullOrEmpty(oUserPrincipal.UserPrincipalName))
                            {                             
                                var upnPrefijoFinal = string.Empty;
                                if (oUserPrincipal.EmailAddress != null)
                                {
                                    var upnPrefijo = oUserPrincipal.EmailAddress;
                                    var startIndex = upnPrefijo.IndexOf("@");
                                    var length = upnPrefijo.Length - startIndex;
                                    upnPrefijoFinal = upnPrefijo.Substring(startIndex, length);
                                }

                                var infoString = string.Empty;
                                if (generarInfo.Equals("S"))
                                {
                                    var info = ((DirectoryEntry)oUserPrincipal.GetUnderlyingObject()).Properties["info"];
                                    if (info.Value != null)
                                    {
                                        infoString = info.Value.ToString();
                                    }
                                }

                                var usuarioAd = new UsuarioAd
                                {
                                    AccountExpirationDate = oUserPrincipal.AccountExpirationDate,
                                    Description = oUserPrincipal.Description,
                                    DisplayName = oUserPrincipal.DisplayName,
                                    DistinguishedName = oUserPrincipal.DistinguishedName,
                                    EmailAddress = oUserPrincipal.EmailAddress,
                                    GivenName = oUserPrincipal.GivenName,
                                    Guid = oUserPrincipal.Guid,
                                    MiddleName = oUserPrincipal.MiddleName,
                                    Name = oUserPrincipal.Name,
                                    SamAccountName = oUserPrincipal.SamAccountName,
                                    Surname = oUserPrincipal.Surname,
                                    Enabled = oUserPrincipal.Enabled,
                                    EstadoCuenta = oUserPrincipal.Enabled != null && oUserPrincipal.Enabled == true ? "Habilitado" : "No habilitado",
                                    UpnPrefijo = upnPrefijoFinal,
                                    InfoString = infoString
                                };
                                listaUser.Add(usuarioAd);
                            }
                        }

                        return listaUser;
                    }
                }
            }
        }

        public string GetAttrOu(string ldap)
        {
            using (var directOu = new DirectoryEntry(ldap, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
            {
                var attrOu = string.Empty;
                if(directOu.Properties["businessCategory"].Value != null)
                {
                    attrOu = Convert.ToString(directOu.Properties["businessCategory"].Value);
                }

                return attrOu;
            };            
        }

        public bool UpdateAttrOu(string ldap, string attr)
        {
            var exito = false;

            using (var directOu = new DirectoryEntry(ldap, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
            {
                directOu.Properties["businessCategory"].Value = attr;
                directOu.CommitChanges();
                directOu.Close();
                exito = true;
                return exito;
            };            
        }

        public string ValidateCodeRouteAd(string codigoRuta)
        {
            using (var directOu = new DirectoryEntry(_sLdapAs, _sUserAdDomain, _sPassAdDomain, AuthenticationTypes.Secure))
            {
                using (var searcher = new DirectorySearcher(directOu))
                {                    
                    searcher.Filter = "(&(objectClass=organizationalUnit)(businessCategory=" + codigoRuta.Trim() + "))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");
                    searcher.PropertiesToLoad.Add("businessCategory");

                    using (var result = searcher.FindAll())
                    {
                        foreach (SearchResult entry in result)
                        {
                            var distinguishedName = entry.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString();
                            var businessCategory = string.Empty;
                            if (entry.GetDirectoryEntry().Properties["businessCategory"].Value != null)
                            {
                                businessCategory = entry.GetDirectoryEntry().Properties["businessCategory"].Value.ToString();
                                if (businessCategory.Equals(codigoRuta))
                                {
                                    return distinguishedName;
                                }
                            }
                        }
                        return string.Empty;
                    };                        
                };
                
            };            
        }

        public List<GrupoAdVm> GetGroupsUserByUserPrincipal(UserPrincipal user)
        {          
            var listGroups = new List<GrupoAdVm>();
            PropertyValueCollection groupsPerson = ((DirectoryEntry)user.GetUnderlyingObject()).Properties["memberOf"];            
            var count = 1;
            foreach (string groupName in groupsPerson)
            {
                using (var objGroup = GetGroup(groupName))
                {
                    if (!(bool)objGroup.IsSecurityGroup)
                    {
                        //var correo = ((DirectoryEntry)objGroup.GetUnderlyingObject()).Properties["mail"];
                        var grupoVm = new GrupoAdVm
                        {
                            NumeroGrupo = count,
                            NombreGrupo = objGroup.Name,
                            UbicacionGrupo = objGroup.DistinguishedName,
                            //CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                            ExisteGrupo = true,
                            DescripcionGrupo = objGroup.Description,
                            TipoGrupo = (bool)objGroup.IsSecurityGroup ? "Grupo Seguridad - " + objGroup.GroupScope.Value : "Grupo Distribución - " + objGroup.GroupScope.Value,
                            Asociado = true
                        };
                        listGroups.Add(grupoVm);
                        count++;
                    }
                };
            }            

            return listGroups;
        }
    }
}
