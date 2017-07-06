
using System;
using System.IO;
using System.Xml.Serialization;

namespace CL.AdmExpertSys.Web.Infrastructure.Helpers
{
    public static class Serializacion
    {
        public static string XmlSerialize(object dataToSerialize)
        {
            if (dataToSerialize == null) return null;

            using (var stringwriter = new StringWriter())
            {
                var serializer = new XmlSerializer(dataToSerialize.GetType());
                serializer.Serialize(stringwriter, dataToSerialize);
                return stringwriter.ToString();
            }
        }

        public static T XmlDeserialize<T>(string xmlText)
        {
            if (String.IsNullOrWhiteSpace(xmlText)) return default(T);

            using (var stringReader = new StringReader(xmlText))
            {
                var serializer = new XmlSerializer(typeof(T));
                return (T)serializer.Deserialize(stringReader);
            }
        }
    }
}
