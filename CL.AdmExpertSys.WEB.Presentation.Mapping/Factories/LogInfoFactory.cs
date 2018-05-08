
using System;
using System.Collections.Generic;
using System.Linq;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class LogInfoFactory
    {
        protected ILogInfoService LogInfoService;

        public LogInfoFactory(
            ILogInfoService logInfoService)
        {
            LogInfoService = logInfoService;
        }

        public List<LogInfoVm> GetAllLogInfo()
        {
            var lista = LogInfoService.FindAll();

            var retorno = (from a in lista
                           select new LogInfoVm
                           {
                               IdInfo = a.IdInfo,
                               MsgInfo = a.MsgInfo,
                               UserInfo = a.UserInfo,
                               FechaInfo = a.FechaInfo,
                               AccionIdInfo = a.AccionIdInfo
                           }).ToList();

            return retorno;
        }

        public LogInfoVm GetLogInfo(long id)
        {
            var tipo = LogInfoService.FindById(id);

            var retorno = new LogInfoVm
            {
                IdInfo = tipo.IdInfo,
                MsgInfo = tipo.MsgInfo,
                UserInfo = tipo.UserInfo,
                FechaInfo = tipo.FechaInfo,
                AccionIdInfo = tipo.AccionIdInfo
            };

            return retorno;
        }

        public void ActualizaLogInfo(LogInfoVm tipo)
        {
            var idInfo = Convert.ToInt32(tipo.IdInfo);
            var tip = LogInfoService.FindById(idInfo);

            tip.MsgInfo = tipo.MsgInfo;
            tip.UserInfo = tipo.UserInfo;
            tip.FechaInfo = tipo.FechaInfo;
            tip.AccionIdInfo = tipo.AccionIdInfo;

            LogInfoService.Update(tip);
        }

        public void CrearLogInfo(LogInfoVm tipo)
        {
            var tip = new LOG_INFO
            {
                MsgInfo = tipo.MsgInfo,
                UserInfo = tipo.UserInfo,
                FechaInfo = tipo.FechaInfo,
                AccionIdInfo = tipo.AccionIdInfo                
            };

            LogInfoService.Create(tip);
        }
    }
}
