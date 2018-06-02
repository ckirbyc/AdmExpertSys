using AutoMapper;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using Pragma.Commons.Data.Patterns.Specification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class UsuarioFactory
    {
        protected IUsuarioService UsuarioService;
        protected ITransversalService TransversalService;

        public UsuarioFactory(
            IUsuarioService usuarioService,
            ITransversalService transversalService)
        {
            UsuarioService = usuarioService;
            TransversalService = transversalService;
        }

        public IList<UsuarioPerfilVm> GetAllUsuarioPerfil()
        {
            var espec = new DirectSpecification<USUARIO>(x => x.Visible);
            var estadoBd = UsuarioService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<USUARIO>, IList<UsuarioPerfilVm>>(estadoBd);
            return vm.ToList();
        }

        public UsuarioPerfilVm GetCreateView()
        {
            var objVm = new UsuarioPerfilVm
            {
                PerfilSelectListItem = TransversalService.GetSelectPerfil()
            };
            return objVm;
        }

        public bool CrearUsuarioPerfil(UsuarioPerfilVm objVm)
        {
            try
            {
                var objBd = new USUARIO
                {
                    Usuario1 = objVm.Usuario1,
                    PerfilId = objVm.PerfilId,
                    EsAdm = objVm.EsAdm,
                    Visible = objVm.Visible
                };
                UsuarioService.Create(objBd);
                return true;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public bool ValidarExistenciaUsuarioBd(string usuario)
        {
            try
            {
                var espec = new DirectSpecification<USUARIO>(x => x.Usuario1 == usuario.ToLower().Trim());
                var estadoBd = UsuarioService.AllMatching(espec);

                if (estadoBd.Any())
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return true;
            }
        }

        public UsuarioPerfilVm GetUsuarioPerfilByNombreCta(string usuario)
        {
            var espec = new DirectSpecification<USUARIO>(x => x.Usuario1 == usuario.ToLower().Trim());
            var objBd = UsuarioService.AllMatching(espec);
            var objVm = new UsuarioPerfilVm();
            if (objBd != null)
            {
                objVm = Mapper.Map(objBd.FirstOrDefault(), objVm);                
            }          

            return objVm;
        }

        public UsuarioPerfilVm GetUsuarioPerfil(decimal id)
        {
            var espec = new DirectSpecification<USUARIO>(x => x.Id == id);
            var objBd = UsuarioService.AllMatching(espec);

            var objVm = Mapper.Map(objBd.FirstOrDefault(), new UsuarioPerfilVm());
            objVm.PerfilSelectListItem = TransversalService.GetSelectPerfil();            

            return objVm;
        }

        public void ActualizaPerfilUsuario(UsuarioPerfilVm tipo)
        {
            var espec = new DirectSpecification<USUARIO>(x => x.Id == tipo.Id);
            var objBd = UsuarioService.AllMatching(espec).FirstOrDefault();

            if (objBd != null)
            {
                objBd.Usuario1 = tipo.Usuario1;
                objBd.PerfilId = tipo.PerfilId;
                objBd.EsAdm = tipo.EsAdm;
                objBd.Visible = tipo.Visible;

                UsuarioService.Update(objBd);
            }
        }

        public void EliminaPerfilUsuario(UsuarioPerfilVm tipo)
        {
            var espec = new DirectSpecification<USUARIO>(x => x.Id == tipo.Id);
            var objBd = UsuarioService.AllMatching(espec).FirstOrDefault();

            if (objBd != null)
            {              
                UsuarioService.Remove(objBd);
            }
        }
    }
}
