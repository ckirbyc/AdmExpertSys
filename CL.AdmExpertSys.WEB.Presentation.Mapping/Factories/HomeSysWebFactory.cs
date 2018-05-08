
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.ADClassLib;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Application.OfficeOnlineClassLib;
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class HomeSysWebFactory
    {       
        protected IHomeSysWebService HomeSysWebService;
        protected Common CommonFactory;
        protected AdLib AdFactory;
        protected Office365 O365Factory;

        public HomeSysWebFactory(
            IHomeSysWebService homeSysWebService)
        {
            HomeSysWebService = homeSysWebService;
        }

        public HomeSysWebFactory()
        {
        }

        public HomeSysWebVm ObtenerVistaHomeSysWeb()
        {
            var homeSys = new HomeSysWebVm { Licencias = ObtenerTipoLicencias(), Ous = ObtenerTipoOus(), UpnPrefijoLista = ObtenerUpnPrefijo(), ListaAccountSkus = GetLicenciasDisponibles()};
            return homeSys;
        }

        private static List<SelectListItem> ObtenerTipoLicencias()
        {
            try
            {
                var office365 = new Office365();
                var listaLicencia = office365.Licenses();
                return listaLicencia.Select(item => new SelectListItem { Text = item.Value, Value = item.Key }).ToList();
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        private static List<SelectListItem> ObtenerTipoOus()
        {
            try
            {
                var adLib = new AdLib();
                var ous = adLib.OUs();
                return ous.Select(item => new SelectListItem { Text = item.Value, Value = item.Key }).ToList();
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        private static List<SelectListItem> ObtenerUpnPrefijo()
        {
            var adLib = new AdLib();
            var upn = adLib.GetUpnPrefijo();
            return upn.Select(item => new SelectListItem { Text = item.Value, Value = item.Key }).ToList();
        }

        public UsuarioAd ObtenerUsuarioExistente(string nombreUsuario)
        {
            try
            {
                var adLib = new AdLib();
                CommonFactory = new Common();
                var userPrincipal = adLib.IsUserExisting(nombreUsuario);
                if (userPrincipal != null)
                {
                    var distinguishedName = userPrincipal.DistinguishedName;                    
                    int startIndex = distinguishedName.IndexOf("OU=");
                    int length = distinguishedName.Length - startIndex;
                    var dnNew = distinguishedName.Substring(startIndex, length);                    
                    var nuevaUbicacion = CommonFactory.GetAppSetting("LdapServidor") + dnNew;
                    userPrincipal.DistinguishedName = nuevaUbicacion;

                    var upnPrefijo = userPrincipal.EmailAddress;
                    startIndex = upnPrefijo.IndexOf("@");
                    length = upnPrefijo.Length - startIndex;
                    userPrincipal.UpnPrefijo = upnPrefijo.Substring(startIndex, length);

                    return userPrincipal;
                }
                return null;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw;
            }
        }
        /// <summary>
        /// Crear usuario en el AD, Sincronizar y Asignar Licencias de O365
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool CrearUsuario(HomeSysWebVm model)
        {
            try
            {
                CommonFactory = new Common();
                AdFactory = new AdLib();

                var nombres = CommonFactory.UppercaseWords(model.Nombres.Trim().ToLower());
                var apellidos = CommonFactory.UppercaseWords(model.Apellidos.Trim().ToLower());
                var username = model.NombreUsuario.ToLower().Trim();
                var pwd = model.Clave.Trim();
                var descripcion = model.Descripcion.Trim();
                //var sTenantName = CommonFactory.GetAppSetting("TenantName");
                //var sPublicDomain = CommonFactory.GetAppSetting("PublicDomain");
                //var sLicense = sTenantName + ":" + model.Licencia.Trim();
                //var sUpn = username + model.UpnPrefijo;

                //Creo el usuario en AD
                if (model.ExisteUsuario)
                {
                    return true;
                }
                using (
                    UserPrincipal sUserPrincipal = AdFactory.CreateNewUser(model.PatchOu, username, pwd, nombres,
                        apellidos, model.UpnPrefijo, pwd, model.ExisteUsuario, descripcion))
                {
                    //Si el usuario se creó correctamente, continuo con DirSync y Asignacion de licencia
                    if (sUserPrincipal != null)
                    {
                        return true;
                    }
                    throw new ArgumentException("Error al crear usuario en AD, intente más tarde.");
                }
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
        }

        public List<string> GetListadoPropiedadesUserAd()
        {
            AdFactory = new AdLib();
            return AdFactory.ObtenerPropiedadesUsuariosAd();
        }

        public static EstructuraArbolAd GetArquitecturaArbolAd()
        {
            try
            {
                var adFactory = new AdLib();
                var listaUser = adFactory.EnumerateOu();
                return HelpFactory.GenerarArbolAdOu(listaUser);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        private static List<MsolAccountSku> GetLicenciasDisponibles()
        {
            try
            {
                var office365 = new Office365();
                return office365.ObtenerMsolAccountSku();
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }
        /// <summary>
        /// Sincroniza Usuarios AD al O365
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool ForzarDirSync(HomeSysWebVm model)
        {
            try
            {
                O365Factory = new Office365();
                var procSync = O365Factory.ForzarDirSync();
                //Task.Delay(TimeSpan.FromSeconds(60)).Wait();
                return procSync;
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
        }

        public void ForzarDirSync()
        {
            try
            {
                O365Factory = new Office365();
                var procSync = O365Factory.ForzarDirSync();                               
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
        }

        /// <summary>
        /// Asigna Licencia Usuiario
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool AsignarLicenciaUsuario(HomeSysWebVm model)
        {
            try
            {
                O365Factory = new Office365();
                var userNameOnline = model.NombreUsuario.ToLower().Trim() + model.UpnPrefijo;
                //Validar que exista usuario en O365
                string sMessage;
                var existeUsr = O365Factory.UserExists(userNameOnline, out sMessage);
                if (existeUsr == false)
                {
                    throw new ArgumentException(sMessage);
                }

                var username = model.NombreUsuario.ToLower().Trim();
                var sUpn = username + model.UpnPrefijo;

                string sMessage2;
                if (O365Factory.AllocateLicense(sUpn, model.Licencia, out sMessage2))
                {
                    return true;
                }
                throw new ArgumentException(sMessage2);
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException(ex.Message);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public bool AsignarLicenciaUsuario(string userNameOnline, string codigoLicencia)
        {
            try
            {
               O365Factory = new Office365();                                               

                string sMessage;
                if (O365Factory.AllocateLicense(userNameOnline, codigoLicencia, out sMessage))
                {
                    return true;
                }
                throw new ArgumentException(sMessage);
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        /// <summary>
        /// Valida credenciales de logeo del usuario en el AD
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="clave"></param>
        /// <returns></returns>
        public bool ValidarCredencialesUsuarioAd(string usuario, string clave)
        {
            try
            {
                AdFactory = new AdLib();
                return AdFactory.ValidateCredentials(usuario, clave);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al validar usuario : " + ex.Message);
            }
        }
        /// <summary>
        /// Factory Deshabilitar cuenta Usuario AD
        /// </summary>
        /// <param name="usuario"></param>
        /// <returns></returns>
        public bool DeshabilitarUsuarioAd(string usuario)
        {
            try
            {
                AdFactory = new AdLib();                               
                AdFactory.DisableUserAccount(usuario);
                //O365Factory = new Office365();
                //O365Factory.ForzarDirSync();
                return true;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al deshabilitar usuario : " + ex.Message);
            }
        }

        public bool DeshabilitarComputadorAd(string computador)
        {
            try
            {
                AdFactory = new AdLib();
                var exito = AdFactory.DisableComputerAccount(computador);
                
                return exito;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al deshabilitar computador : " + ex.Message);
            }
        }

        /// <summary>
        /// Obtiene Grupo del AD por Nombre
        /// </summary>
        /// <param name="nomGrupoSel"></param>
        /// <returns></returns>
        public List<GrupoAdVm> ObtenerGrupoAdByNombre(string nomGrupoSel)
        {
            try
            {
                var listaGrupo = new List<GrupoAdVm>();
                AdFactory = new AdLib();

                var i = 1;
                foreach (var nomGrupo in nomGrupoSel.Split(','))
                {
                    if (!string.IsNullOrEmpty(nomGrupo))
                    {
                        using (var objGrupo = AdFactory.GetGroup(nomGrupo))
                        {
                            if (objGrupo != null)
                            {
                                DirectorySearcher objSearch;
                                using (var objDirecEnt = (DirectoryEntry) objGrupo.GetUnderlyingObject())
                                {
                                    objSearch = new DirectorySearcher(objDirecEnt);
                                }
                                objSearch.PropertiesToLoad.Add("mail");
                                ResultPropertyCollection rpc;
                                using (SearchResultCollection results = objSearch.FindAll())
                                {
                                    rpc = results[0].Properties;
                                }
                                var correoGrupo = rpc["mail"][0].ToString();

                                var grupoVm = new GrupoAdVm
                                {
                                    NumeroGrupo = i,
                                    NombreGrupo = objGrupo.DisplayName,
                                    UbicacionGrupo = objGrupo.DistinguishedName,
                                    CorreoGrupo = correoGrupo,
                                    ExisteGrupo = true
                                };
                                listaGrupo.Add(grupoVm);
                                i++;
                            }
                            //else
                            //{
                            //    var grupoVm = new GrupoAdVm
                            //    {
                            //        NumeroGrupo = i,
                            //        NombreGrupo = nomGrupo,
                            //        UbicacionGrupo = string.Empty,
                            //        ExisteGrupo = false
                            //    };
                            //    listaGrupo.Add(grupoVm);
                            //}
                        }
                    }
                    //i++;
                }

              return listaGrupo;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al obtener grupo : " + ex.Message);
            }
        }

        public List<GrupoAdVm> ObtenerListadoGrupoAdByOu(string sOu)
        {
            try
            {
                var listaGrupo = new List<GrupoAdVm>();
                AdFactory = new AdLib();

                var objListaGroup = AdFactory.GetListGroupByOu(sOu);

                var i = 1;
                foreach (GroupPrincipal objGroup in objListaGroup.FindAll())
                {                                        
                    var correo = ((DirectoryEntry)objGroup.GetUnderlyingObject()).Properties["mail"];                    

                    var grupoVm = new GrupoAdVm
                    {
                        NumeroGrupo = i,
                        NombreGrupo = objGroup.Name,
                        UbicacionGrupo = objGroup.DistinguishedName,
                        CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                        ExisteGrupo = true,
                        DescripcionGrupo = objGroup.Description,
                        TipoGrupo = (bool)objGroup.IsSecurityGroup ? "Grupo Seguridad - " + objGroup.GroupScope.Value : "Grupo Distribución - " + objGroup.GroupScope.Value
                    };
                    listaGrupo.Add(grupoVm);
                    i++;                    
                }

                objListaGroup.Dispose();

                return listaGrupo;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al obtener grupo : " + ex.Message);
            }
        }

        public List<GrupoAdVm> ObtenerListadoGrupoAdByOu(string sOu, string adUsr)
        {
            try
            {
                var listaGrupo = new List<GrupoAdVm>();
                AdFactory = new AdLib();

                var userAd = AdFactory.GetUser(adUsr);

                var objListaGroup = AdFactory.GetListGroupByOu(sOu);

                var i = 1;                
                foreach (GroupPrincipal objGroup in objListaGroup.FindAll())
                {
                    var correo = ((DirectoryEntry)objGroup.GetUnderlyingObject()).Properties["mail"];
                    var asocUsrGrp = AdFactory.IsUserGroupMember(userAd, objGroup);                                      

                    var grupoVm = new GrupoAdVm
                    {
                        NumeroGrupo = i,
                        NombreGrupo = objGroup.Name,
                        UbicacionGrupo = objGroup.DistinguishedName,
                        CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                        ExisteGrupo = true,
                        DescripcionGrupo = objGroup.Description,
                        TipoGrupo = (bool)objGroup.IsSecurityGroup ? "Grupo Seguridad - " + objGroup.GroupScope.Value : "Grupo Distribución - " + objGroup.GroupScope.Value,
                        Asociado = asocUsrGrp
                    };
                    listaGrupo.Add(grupoVm);
                    i++;
                }

                userAd.Dispose();
                objListaGroup.Dispose();

                return listaGrupo;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al obtener grupo : " + ex.Message);
            }
        }

        public bool GuardarGrupos(string nomGrupos, string userName)
        {
            try
            {
                AdFactory = new AdLib();
                O365Factory = new Office365();                
               
                var exito = false;
                foreach (var nomGrupo in nomGrupos.Split(','))
                {
                    if (!string.IsNullOrEmpty(nomGrupo))
                    {
                        exito = AdFactory.AddUserToGroup(userName, nomGrupo);
                    }
                }
                O365Factory.ForzarDirSync();
                return exito;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al guardar grupo : " + ex.Message);
            }
        }

        public bool AsociarGrupoUsuario(string nomGrupo, string userName)
        {
            try
            {
                AdFactory = new AdLib();
                var exito = AdFactory.AddUserToGroup(userName, nomGrupo);               
                return exito;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al asociar grupo a usuario : " + ex.Message);
            }
        }

        public bool DesAsociarGrupoUsuario(string nomGrupo, string userName)
        {
            try
            {
                AdFactory = new AdLib();
                var exito = AdFactory.RemoveUserFromGroup(userName, nomGrupo);
                return exito;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al desasociar grupo a usuario : " + ex.Message);
            }
        }

        public bool ExisteLicenciaUsuarioPortal(string userName)
        {
            try
            {
                string msgError;
                O365Factory = new Office365();
                var existeLicencia = O365Factory.IsLicensedUser(userName, out msgError);
                return existeLicencia;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public bool ExisteUsuarioPortal(string userName)
        {
            try
            {
                
                O365Factory = new Office365();                
                //Validar que exista usuario en O365
                string sMessage;
                var existeUsr = O365Factory.UserExists(userName, out sMessage);
                                
                return existeUsr;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public bool EliminarCuentaAd(string userName)
        {
            try
            {
                AdFactory = new AdLib();
                var exito = AdFactory.DeleteUser(userName);
                return exito;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public GrupoAdVm GetCreateViewGrupo()
        {
            var objVm = new GrupoAdVm();
            return objVm;
        }

        public void CrearGrupoDistribucion(GrupoAdVm model)
        {
            AdFactory = new AdLib();
            AdFactory.CreateNewGroup(model.UbicacionGrupo, model.NombreGrupo, model.DescripcionGrupo, model.CorreoGrupo, GroupScope.Global, false);
        }

        public void ActualizarGrupoDistribucion(GrupoAdVm model)
        {
            AdFactory = new AdLib();
            AdFactory.UpdateGroup(model.UbicacionGrupo, model.NombreGrupo, model.NombreGrupoAnterior, model.DescripcionGrupo, model.CorreoGrupo, GroupScope.Global, false);
        }

        public GrupoAdVm GetGrupoDistribucion(string nombreGrupo)
        {
            
            AdFactory = new AdLib();
            CommonFactory = new Common();
            var objAd = AdFactory.GetGroup(nombreGrupo);

            var correo = ((DirectoryEntry)objAd.GetUnderlyingObject()).Properties["mail"];
            var replaceUbicacion = @"CN=" + objAd.Name + ",";
            var nuevaUbicacion = CommonFactory.GetAppSetting("LdapServidor") + objAd.DistinguishedName.Replace(replaceUbicacion,"");

            var objVm = new GrupoAdVm
            {
                NumeroGrupo = 1,
                NombreGrupo = objAd.Name,
                NombreGrupoAnterior = objAd.Name,
                UbicacionGrupo = nuevaUbicacion,
                CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                ExisteGrupo = true,
                DescripcionGrupo = objAd.Description,
                TipoGrupo = (bool)objAd.IsSecurityGroup ? "Grupo Seguridad - " + objAd.GroupScope.Value : "Grupo Distribución - " + objAd.GroupScope.Value
            };

            return objVm;
        }

        public void EliminarGrupoDistribucion(GrupoAdVm model)
        {
            AdFactory = new AdLib();
            AdFactory.DeleteGroup(model.NombreGrupo);
        }

        public bool ActualizarCuentaUsuario(HomeSysWebVm model)
        {
            AdFactory = new AdLib();            
            try
            {
                //Actualizar cuenta usuario en el AD
                var exitoActUsr = AdFactory.UpdateUser(model);                                
                return exitoActUsr;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }
    }
}
