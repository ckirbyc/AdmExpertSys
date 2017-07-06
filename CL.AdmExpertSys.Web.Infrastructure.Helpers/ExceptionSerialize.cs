
using System;

namespace CL.AdmExpertSys.Web.Infrastructure.Helpers
{
    [Serializable]
    public class ExceptionSerialize
    {
        public string HelpLink { get; set; }
        public int HResult { get; set; }
        public ExceptionSerialize InnerException { get; set; }
        public string Message { get; set; }
        public string Source { get; set; }
        public string StackTrace { get; set; }

        public ExceptionSerialize() { }
        public ExceptionSerialize(Exception exception)
        {
            CloneException(this, exception);
        }

        private void CloneException(ExceptionSerialize objeto, Exception exception)
        {
            HelpLink = exception.HelpLink;
            HResult = exception.HResult;
            Message = exception.Message;
            Source = exception.Source;
            StackTrace = exception.StackTrace;

            if (exception.InnerException != null)
            {
                CloneException(objeto, exception.InnerException);
            }
        }
    }
}
