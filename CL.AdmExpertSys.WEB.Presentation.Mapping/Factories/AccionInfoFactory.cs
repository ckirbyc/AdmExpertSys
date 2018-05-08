
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System.Collections.Generic;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class AccionInfoFactory
    {
        protected IAccionInfoService AccionInfoService;

        public AccionInfoFactory(
            IAccionInfoService accionInfoService)
        {
            AccionInfoService = accionInfoService;
        }

        public List<AccionInfoVm> GetAllAccionInfo()
        {
            var lista = AccionInfoService.FindAll();

            var retorno = (from a in lista
                           select new AccionInfoVm
                           {
                               IdAccionInfo = a.IdAccionInfo,
                               NombreAccionInfo = a.NombreAccionInfo
                           }).ToList();

            return retorno;
        }

        public AccionInfoVm GetAccionInfo(long id)
        {
            var tipo = AccionInfoService.FindById(id);

            var retorno = new AccionInfoVm
            {
                IdAccionInfo = tipo.IdAccionInfo,
                NombreAccionInfo = tipo.NombreAccionInfo
            };

            return retorno;
        }

        public void ActualizaAccionInfo(AccionInfoVm tipo)
        {
            var tip = AccionInfoService.FindById(tipo.IdAccionInfo);

            tip.NombreAccionInfo = tipo.NombreAccionInfo;

            AccionInfoService.Update(tip);
        }

        public void CrearAccionInfo(AccionInfoVm tipo)
        {
            var tip = new ACCION_INFO
            {
                NombreAccionInfo = tipo.NombreAccionInfo
            };

            AccionInfoService.Create(tip);
        }
    }
}
