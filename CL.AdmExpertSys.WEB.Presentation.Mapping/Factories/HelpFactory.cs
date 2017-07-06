
using System.Linq;
using System.Text;
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using System;
using System.Collections.Generic;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    [Serializable]
    public class HelpFactory
    {
        public static EstructuraArbolAd GenerarArbolAdOu(List<Ou> listaOu)
        {
            var estructura = new EstructuraArbolAd {Cabecera = @"Empresas Agrosuper Sap"};
            var codHtml = new StringBuilder(string.Empty);
            var idPublicado = new List<int>();

            foreach (var ou in listaOu)
            {
                var tieneHijo = listaOu.Count(x => x.IdPadreOu == ou.IdOu);
                if (tieneHijo == 0)
                {
                    if (idPublicado.Count(x => x == ou.IdOu) == 0 && ou.Nivel == 1)
                    {
                        codHtml.Append("<li id='" + ou.Ldap + "'>" + ou.Nombre.Replace("OU=", "") + "</li>");
                    }
                    idPublicado.Add(ou.IdOu);
                }
                else
                {
                    if (idPublicado.Count(x => x == ou.IdOu) == 0)
                    {
                        if(ou.Nivel == 1)
                            codHtml.Append("<li id='" + ou.Ldap + "'>" + ou.Nombre.Replace("OU=", ""));
                        idPublicado.Add(ou.IdOu);
                        var listaOuHijo = listaOu.Where(x => x.IdPadreOu == ou.IdOu && x.Nivel == 2).ToList();

                        foreach (var ou1 in listaOuHijo)
                        {
                            if (idPublicado.Count(x => x == ou1.IdOu) == 0)
                            {
                                var listaOuHijo2 = listaOu.Where(x => x.IdPadreOu == ou1.IdOu && x.Nivel == 3).ToList();
                                if (!listaOuHijo2.Any())
                                {
                                    codHtml.Append("<ul>");
                                    codHtml.Append("<li id='" + ou1.Ldap + "'>");
                                    codHtml.Append(ou1.Nombre.Replace("OU=", ""));
                                    codHtml.Append("</li>");
                                    codHtml.Append("</ul>");
                                }
                                else
                                {
                                    codHtml.Append("<ul>");
                                    codHtml.Append("<li id='" + ou1.Ldap + "'>");
                                    codHtml.Append(ou1.Nombre.Replace("OU=", ""));
                                    foreach (var ou2 in listaOuHijo2)
                                    {
                                        var listaOuHijo3 = listaOu.Where(x => x.IdPadreOu == ou2.IdOu && x.Nivel == 4).ToList();
                                        if (!listaOuHijo3.Any())
                                        {
                                            codHtml.Append("<ul>");
                                            codHtml.Append("<li id='" + ou2.Ldap + "'>");
                                            codHtml.Append(ou2.Nombre.Replace("OU=", ""));
                                            codHtml.Append("</li>");
                                            codHtml.Append("</ul>");
                                        }
                                        else
                                        {
                                            codHtml.Append("<ul>");
                                            codHtml.Append("<li id='" + ou2.Ldap + "'>");
                                            codHtml.Append(ou2.Nombre.Replace("OU=", ""));
                                            foreach (var ou3 in listaOuHijo3)
                                            {
                                                var listaOuHijo4 = listaOu.Where(x => x.IdPadreOu == ou3.IdOu && x.Nivel == 5).ToList();
                                                if (!listaOuHijo4.Any())
                                                {
                                                    codHtml.Append("<ul>");
                                                    codHtml.Append("<li id='" + ou3.Ldap + "'>");
                                                    codHtml.Append(ou3.Nombre.Replace("OU=", ""));
                                                    codHtml.Append("</li>");
                                                    codHtml.Append("</ul>");
                                                }
                                                else
                                                {
                                                    codHtml.Append("<ul>");
                                                    codHtml.Append("<li id='" + ou3.Ldap + "'>");
                                                    codHtml.Append(ou3.Nombre.Replace("OU=", ""));
                                                    foreach (var ou4 in listaOuHijo4)
                                                    {
                                                        var listaOuHijo5 = listaOu.Where(x => x.IdPadreOu == ou4.IdOu && x.Nivel == 6).ToList();
                                                        if (!listaOuHijo5.Any())
                                                        {
                                                            codHtml.Append("<ul>");
                                                            codHtml.Append("<li id='" + ou4.Ldap + "'>");
                                                            codHtml.Append(ou4.Nombre.Replace("OU=", ""));
                                                            codHtml.Append("</li>");
                                                            codHtml.Append("</ul>");
                                                        }
                                                        else
                                                        {
                                                            codHtml.Append("<ul>");
                                                            codHtml.Append("<li id='" + ou4.Ldap + "'>");
                                                            codHtml.Append(ou4.Nombre.Replace("OU=", ""));
                                                            foreach (var ou5 in listaOuHijo5)
                                                            {
                                                                var listaOuHijo6 = listaOu.Where(x => x.IdPadreOu == ou5.IdOu && x.Nivel == 7).ToList();
                                                                if (!listaOuHijo6.Any())
                                                                {
                                                                    codHtml.Append("<ul>");
                                                                    codHtml.Append("<li id='" + ou5.Ldap + "'>");
                                                                    codHtml.Append(ou5.Nombre.Replace("OU=", ""));
                                                                    codHtml.Append("</li>");
                                                                    codHtml.Append("</ul>");
                                                                }
                                                                else
                                                                {
                                                                    codHtml.Append("<ul>");
                                                                    codHtml.Append("<li id='" + ou5.Ldap + "'>");
                                                                    codHtml.Append(ou5.Nombre.Replace("OU=", ""));
                                                                    foreach (var ou6 in listaOuHijo6)
                                                                    {
                                                                        codHtml.Append("<ul>");
                                                                        codHtml.Append("<li id='" + ou6.Ldap + "'>");
                                                                        codHtml.Append(ou6.Nombre.Replace("OU=", ""));
                                                                        codHtml.Append("</li>");
                                                                        codHtml.Append("</ul>");
                                                                    }
                                                                    codHtml.Append("</li>");
                                                                    codHtml.Append("</ul>");
                                                                }
                                                            }
                                                            codHtml.Append("</li>");
                                                            codHtml.Append("</ul>");
                                                        }
                                                    }
                                                    codHtml.Append("</li>");
                                                    codHtml.Append("</ul>");
                                                }
                                            }
                                            codHtml.Append("</li>");
                                            codHtml.Append("</ul>");
                                        }
                                    }
                                    codHtml.Append("</li>");
                                    codHtml.Append("</ul>");
                                }
                            }
                            idPublicado.Add(ou1.IdOu);
                        }
                        if(ou.Nivel == 1)
                            codHtml.Append("</li>");
                    }
                    idPublicado.Add(ou.IdOu);
                }
                
            }
            estructura.CodArbolHtml = codHtml.ToString();
            return estructura;
        }
    }
}
