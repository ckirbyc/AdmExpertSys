
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.ADClassLib;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Application.OfficeOnlineClassLib;
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using ClosedXML.Excel;
using System;
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
        protected IMantenedorLicenciaService MantenedorLicenciaService;

        public HomeSysWebFactory(
            IHomeSysWebService homeSysWebService,
            IMantenedorLicenciaService mantenedorLicenciaService)
        {
            HomeSysWebService = homeSysWebService;
            MantenedorLicenciaService = mantenedorLicenciaService;
        }

        public HomeSysWebFactory()
        {
        }

        public HomeSysWebVm ObtenerVistaHomeSysWeb()
        {
            var homeSys = new HomeSysWebVm {
                Licencias = ObtenerTipoLicencias(),
                Ous = ObtenerTipoOus(),
                UpnPrefijoLista = ObtenerUpnPrefijo(),
                ListaAccountSkus = GetLicenciasDisponibles(),
                CodigoLicenciaLista = ObtenerSelectMantenedorLicencia()
            };
            return homeSys;
        }

        private List<SelectListItem> ObtenerSelectMantenedorLicencia()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    var lista = (from a in entityContext.MANTENEDOR_LICENCIA.Where(x => x.Vigente).ToList()
                                 orderby a.ROL_CARGO.Nombre
                                 select a).AsEnumerable()
                    .Select(x => new SelectListItem
                    {
                        Value = x.Codigo,
                        Text = x.ROL_CARGO.Nombre
                    }).ToList();

                    return lista;
                }                
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
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

                //Creo el usuario en AD
                if (model.ExisteUsuario)
                {
                    return true;
                }
                using (
                    UserPrincipal sUserPrincipal = AdFactory.CreateNewUser(model.PatchOu, username, pwd, nombres,
                        apellidos, model.UpnPrefijo, pwd, model.ExisteUsuario, descripcion, model.Info))
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
            O365Factory = new Office365();                                               

            string sMessage;
            if (O365Factory.AllocateLicense(userNameOnline, codigoLicencia, out sMessage))
            {
                return true;
            }
            return false;            
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
                return false;
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
                foreach (GroupPrincipal objGroup in objListaGroup.ToList())
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
                    objGroup.Dispose();
                    i++;                    
                }                

                return listaGrupo;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new ArgumentException("Error al obtener grupo : " + ex.Message);
            }
        }

        public List<GrupoAdVm> ObtenerListadoGrupoAdByOuMantGroup(string sOu)
        {
            try
            {
                var listaGrupo = new List<GrupoAdVm>();
                AdFactory = new AdLib();

                var objListaGroup = AdFactory.GetListGroupByOuMantGroup(sOu);

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
                    objGroup.Dispose();
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

                var j = 1;
                foreach (GroupPrincipal objGroup in userAd.GetGroups())
                {
                    if (!(bool)objGroup.IsSecurityGroup)
                    {
                        var correo = ((DirectoryEntry)objGroup.GetUnderlyingObject()).Properties["mail"];
                        var grupoVm = new GrupoAdVm
                        {
                            NumeroGrupo = j,
                            NombreGrupo = objGroup.Name,
                            UbicacionGrupo = objGroup.DistinguishedName,
                            CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                            ExisteGrupo = true,
                            DescripcionGrupo = objGroup.Description,
                            TipoGrupo = (bool)objGroup.IsSecurityGroup ? "Grupo Seguridad - " + objGroup.GroupScope.Value : "Grupo Distribución - " + objGroup.GroupScope.Value,
                            Asociado = true
                        };
                        listaGrupo.Add(grupoVm);
                        j++;
                    }
                    objGroup.Dispose();
                }

                var objListaGroup = AdFactory.GetListGroupByOu(sOu);

                var i = j;                
                foreach (GroupPrincipal objGroup in objListaGroup.ToList())
                {                                       
                    var asocUsrGrp = AdFactory.IsUserGroupMember(userAd, objGroup);
                    if (!asocUsrGrp)
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
                            TipoGrupo = (bool)objGroup.IsSecurityGroup ? "Grupo Seguridad - " + objGroup.GroupScope.Value : "Grupo Distribución - " + objGroup.GroupScope.Value,
                            Asociado = asocUsrGrp
                        };
                        listaGrupo.Add(grupoVm);
                        i++;
                    }                    
                    objGroup.Dispose();                    
                }

                userAd.Dispose();                

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
            string msgError;
            O365Factory = new Office365();
            var existeLicencia = O365Factory.IsLicensedUser(userName, out msgError);
            return existeLicencia;            
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

            //Obtener usuarios del grupo
            var listaUsuario = new List<UsuarioAd>();
            foreach (var userObj in objAd.GetMembers().ToList())
            {
                if (userObj.UserPrincipalName != null)
                {
                    var correoUsrObj = ((DirectoryEntry)userObj.GetUnderlyingObject()).Properties["mail"];
                    var correoUsr = correoUsrObj.Value != null ? correoUsrObj.Value.ToString() : string.Empty;
                    var objUsr = new UsuarioAd
                    {
                        DisplayName = userObj.DisplayName,
                        Description = userObj.Description,
                        DistinguishedName = userObj.DistinguishedName,
                        SamAccountName = userObj.SamAccountName,
                        Name = userObj.Name,
                        EmailAddress = correoUsr
                    };
                    listaUsuario.Add(objUsr);
                }                
                userObj.Dispose();
            }

            var correo = ((DirectoryEntry)objAd.GetUnderlyingObject()).Properties["mail"];
            var replaceUbicacion = @"CN=" + objAd.Name + ",";
            var nuevaUbicacion = CommonFactory.GetAppSetting("LdapServidor") + objAd.DistinguishedName.Replace(replaceUbicacion, "");

            var objVm = new GrupoAdVm
            {
                NumeroGrupo = 1,
                NombreGrupo = objAd.Name,
                NombreGrupoAnterior = objAd.Name,
                UbicacionGrupo = nuevaUbicacion,
                CorreoGrupo = correo.Value != null ? correo.Value.ToString() : string.Empty,
                ExisteGrupo = true,
                DescripcionGrupo = objAd.Description,
                TipoGrupo = (bool)objAd.IsSecurityGroup ? "Grupo Seguridad - " + objAd.GroupScope.Value : "Grupo Distribución - " + objAd.GroupScope.Value,
                ListaUsuarioAd = listaUsuario
            };

            objAd.Dispose();

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

        public List<UsuarioAd> ObtenerListaCuentaUsuario(string generarInfo)
        {
            var listaAccount = new List<UsuarioAd>();
            try
            {
                AdFactory = new AdLib();                
                listaAccount = AdFactory.GetListAccountUsers(generarInfo);              
                return listaAccount;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return listaAccount;
            }
        }

        public List<UsuarioAd> ObtenerListaCuentaUsuarioCompletoAd(string generarInfo)
        {
            var listaAccount = new List<UsuarioAd>();
            try
            {
                AdFactory = new AdLib();
                listaAccount = AdFactory.GetListAccountAllUsersAd(generarInfo);
                return listaAccount;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return listaAccount;
            }
        }

        public List<UsuarioAd> ObtenerListaCuentaUsuarioLicense()
        {
            var listaAccount = new List<UsuarioAd>();
            try
            {
                AdFactory = new AdLib();
                O365Factory = new Office365();
                listaAccount = AdFactory.GetListAccountUsers("S");

                var sMess = string.Empty;
                listaAccount = O365Factory.GetLicensedUserMassive(listaAccount, out sMess);

                return listaAccount;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return listaAccount;
            }
        }

        public List<UsuarioAd> ObtenerListaCuentaUsuarioLicenseCompletoAd()
        {
            var listaAccount = new List<UsuarioAd>();
            try
            {
                AdFactory = new AdLib();
                O365Factory = new Office365();
                listaAccount = AdFactory.GetListAccountAllUsersAd("S");

                var sMess = string.Empty;
                listaAccount = O365Factory.GetLicensedUserMassive(listaAccount, out sMess);

                return listaAccount;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return listaAccount;
            }
        }

        public XLWorkbook ExportarArchivoExcelReporteCuentaUsuario(string licencia)
        {
            try
            {
                using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
                {
                    using (var hojaRep = workbook.Worksheets.Add("Reporte Cuentas"))
                    {
                        //Imprime Encabezado
                        hojaRep.Cell(1, 1).Value = "Nombre";
                        hojaRep.Cell(1, 1).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 1).Style.Font.Bold = true;

                        hojaRep.Cell(1, 2).Value = "Cuenta";
                        hojaRep.Cell(1, 2).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 2).Style.Font.Bold = true;

                        hojaRep.Cell(1, 3).Value = "Descripción";
                        hojaRep.Cell(1, 3).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 3).Style.Font.Bold = true;

                        hojaRep.Cell(1, 4).Value = "Upn Prefijo";
                        hojaRep.Cell(1, 4).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 4).Style.Font.Bold = true;

                        hojaRep.Cell(1, 5).Value = "Info";
                        hojaRep.Cell(1, 5).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 5).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 5).Style.Font.Bold = true;

                        hojaRep.Cell(1, 6).Value = "Licencia";
                        hojaRep.Cell(1, 6).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 6).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 6).Style.Font.Bold = true;

                        hojaRep.Cell(1, 7).Value = "Ubicación";
                        hojaRep.Cell(1, 7).Style.Font.FontColor = XLColor.White;
                        hojaRep.Cell(1, 7).Style.Fill.BackgroundColor = XLColor.Red;
                        hojaRep.Cell(1, 7).Style.Font.Bold = true;

                        CommonFactory = new Common();
                        var cantNivel = Convert.ToInt64(CommonFactory.GetAppSetting("NivelesArbolAd"));
                        var numColCabecera = 8;

                        for (int i = 1; i <= cantNivel; i++)
                        {
                            var nombreNivelCabecera = @"Nivel " + i.ToString();
                            hojaRep.Cell(1, numColCabecera).Value = nombreNivelCabecera;
                            hojaRep.Cell(1, numColCabecera).Style.Font.FontColor = XLColor.White;
                            hojaRep.Cell(1, numColCabecera).Style.Fill.BackgroundColor = XLColor.Red;
                            hojaRep.Cell(1, numColCabecera).Style.Font.Bold = true;

                            numColCabecera++;
                        }

                        //Obtener los datos para poblar archivo                        
                        var listCtaOrdenada = new List<UsuarioAd>();
                        if (licencia.Equals("N"))
                        {
                            listCtaOrdenada = ObtenerListaCuentaUsuario("S").OrderBy(x => x.DistinguishedName).ToList();
                        }
                        else
                        {
                            listCtaOrdenada = HiloReporteLicencia.GetListaReporteLicencia().OrderBy(x => x.DistinguishedName).ToList();
                        }

                        var listaExcelRepCtaUsr = new List<string[]>(listCtaOrdenada.Count());
                        var cantColArray = 7 + cantNivel;
                        foreach (var ctaUsr in listCtaOrdenada)
                        {
                            var listaRepCtaUsr = new string[cantColArray];
                            listaRepCtaUsr[0] = ctaUsr.Name;
                            listaRepCtaUsr[1] = ctaUsr.SamAccountName;
                            listaRepCtaUsr[2] = ctaUsr.Description;
                            listaRepCtaUsr[3] = ctaUsr.SamAccountName + ctaUsr.UpnPrefijo;
                            listaRepCtaUsr[4] = ctaUsr.InfoString;
                            listaRepCtaUsr[5] = ctaUsr.Licenses;

                            var listaOu = new List<OuExcelVm>();
                            int startIndexOu = ctaUsr.DistinguishedName.IndexOf("OU=");
                            if (startIndexOu > 0)
                            {
                                int lengthOu = ctaUsr.DistinguishedName.Length - startIndexOu;
                                var dnNewOu = ctaUsr.DistinguishedName.Substring(startIndexOu, lengthOu);

                                listaRepCtaUsr[6] = dnNewOu;

                                var dnCompleto = ctaUsr.DistinguishedName;

                                for (int j = 1; j <= cantNivel; j++)
                                {
                                    int startIndex = dnCompleto.IndexOf("OU=");
                                    if (startIndex > 0)
                                    {
                                        int length = dnCompleto.Length - startIndex;
                                        var dnNewAux = dnCompleto.Substring(startIndex, length);
                                        int startlengthAux = dnNewAux.IndexOf(",");
                                        if (startlengthAux > 0)
                                        {
                                            var dnNew = dnNewAux.Substring(0, startlengthAux);
                                            var ouVm = new OuExcelVm
                                            {
                                                Numero = j,
                                                Nombre = dnNew.Replace("OU=", string.Empty)
                                            };
                                            listaOu.Add(ouVm);
                                            dnCompleto = dnCompleto.Replace(dnNew, string.Empty);
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                listaRepCtaUsr[6] = ctaUsr.DistinguishedName;                                
                            }

                            var colArray = 7;
                            foreach (var ouVm in listaOu.OrderByDescending(x => x.Numero))
                            {
                                listaRepCtaUsr[colArray] = ouVm.Nombre;
                                colArray++;
                            }

                            listaExcelRepCtaUsr.Add(listaRepCtaUsr);                            
                        }

                        hojaRep.Cell(2, 1).InsertData(listaExcelRepCtaUsr.AsEnumerable());
                        hojaRep.Columns().AdjustToContents();
                    }
                    return workbook;
                }
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw;
            }            
        }
    }
}
