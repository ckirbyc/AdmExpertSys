
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace CL.AdmExpertSys.Web.Infrastructure.LogTransaccional
{
    /// <summary>
    /// Object Dumper
    /// Lee el grafo de un objeto y lo formatea a string
    /// </summary>
    /// <remarks>
    /// $Revision: 25598 $
    /// $Author: ckirby $
    /// $Date: 2011-12-20 15:51:29 -0300 (Tue, 20 Dec 2011) $
    /// </remarks>
    public class ObjectDumper : IDisposable
    {
        private TextWriter _writer;
        private int _level;
        private readonly int _depth;

        /// <summary>
        /// Escribe
        /// </summary>
        /// <param name="element"></param>
        public static void Write(object element)
        {
            Write(element, 0);
        }

        /// <summary>
        /// Escribe
        /// </summary>
        /// <param name="element"></param>
        /// <param name="depth"></param>
        public static void Write(object element, int depth)
        {
            Write(element, depth, Console.Out);
        }

        /// <summary>
        /// Escribe
        /// </summary>
        /// <param name="element"></param>
        /// <param name="depth"></param>
        /// <param name="log"></param>
        public static void Write(object element, int depth, TextWriter log)
        {
            using (var dumper = new ObjectDumper(depth))
            {
                dumper._writer = log;
                dumper.WriteObject(null, element);
            }
        }

        /// <summary>
        /// Escribe
        /// </summary>
        /// <param name="depth"></param>
        private ObjectDumper(int depth)
        {
            _depth = depth;
        }

        /// <summary>
        /// Escribe
        /// </summary>
        /// <param name="s"></param>
        private void Write(string s)
        {
            if (s != null)
            {
                _writer.Write(s);
            }
        }

        /// <summary>
        /// Escribe indentación
        /// </summary>
        private void WriteIndent()
        {
            for (int i = 0; i < _level; i++) _writer.Write("  ");
        }

        /// <summary>
        /// Escribe una linea
        /// </summary>
        private void WriteLine()
        {
            _writer.WriteLine();
        }

        /*
                private void WriteTab()
                {
                    Write("  ");
                    while (_pos % 8 != 0) Write(" ");
                }
        */

        /// <summary>
        /// Escribe el grafo de un objeto
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="element"></param>
        private void WriteObject(string prefix, object element)
        {
            if (element == null || element is ValueType || element is string)
            {
                WriteIndent();
                Write(prefix);
                WriteValue(element);
                WriteLine();
            }
            else
            {
                var enumerableElement = element as IEnumerable;
                if (enumerableElement != null)
                {
                    EscribirMiembro(prefix, enumerableElement);
                }
                else
                {
                    var members = element.GetType().GetMembers(BindingFlags.Public | BindingFlags.Instance);
                    WriteIndent();
                    Write(prefix);
                    bool propWritten = RecorrerMiembros(element, false, members);
                    if (propWritten) WriteLine();
                    if (_level < _depth)
                    {
                        ProcesarMiembrosNivelAlto(element, members);
                    }
                }
            }
        }

        /// <summary>
        /// Procesa miembros
        /// </summary>
        /// <param name="element"></param>
        /// <param name="members"></param>
        private void ProcesarMiembrosNivelAlto(object element, IEnumerable<MemberInfo> members)
        {
            foreach (var m in members)
            {
                var f = m as FieldInfo;
                var p = m as PropertyInfo;
                if (f != null || p != null)
                {
                    Type t = f != null ? f.FieldType : p.PropertyType;
                    if (!(t.IsValueType || t == typeof(string)))
                    {

                        if (m.ToString().Contains("CL.AdmExpertSys"))
                        {
                            object value = f != null ? f.GetValue(element) : p.GetValue(element, null);
                            if (value != null)
                            {
                                _level++;
                                WriteObject(m.Name + ": ", value);
                                _level--;
                            }
                        }

                    }
                }
            }
        }

        /// <summary>
        /// Recorre Miembros
        /// </summary>
        /// <param name="element"></param>
        /// <param name="propWritten"></param>
        /// <param name="members"></param>
        /// <returns></returns>
        private bool RecorrerMiembros(object element, bool propWritten, IEnumerable<MemberInfo> members)
        {
            foreach (MemberInfo m in members)
            {
                var f = m as FieldInfo;
                var p = m as PropertyInfo;
                if (f != null || p != null)
                {
                    if (propWritten)
                    {
                        //WriteTab();
                        WriteLine();
                    }
                    else
                    {
                        propWritten = true;
                    }
                    Write(m.Name);
                    Write(" = ");
                    var t = f != null ? f.FieldType : p.PropertyType;
                    if (t.IsValueType || t == typeof(string))
                    {

                        if (f != null)
                        {
                            WriteValue(f.GetValue(element));
                            Write("; ");
                        }
                        else
                        {

                            //if (p.GetType().IsGenericParameter)
                            //{
                            WriteValue(p.GetValue(element, null));
                            Write("; ");
                            //}
                        }

                    }
                    else
                    {
                        Write(typeof(IEnumerable).IsAssignableFrom(t) ? "..." : "{ }");
                    }
                }
            }
            return propWritten;
        }

        /// <summary>
        /// Escribe un miembro
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="enumerableElement"></param>
        private void EscribirMiembro(string prefix, IEnumerable enumerableElement)
        {
            foreach (object item in enumerableElement)
            {
                if (item is IEnumerable && !(item is string))
                {
                    WriteIndent();
                    Write(prefix);
                    Write("...");
                    WriteLine();
                    if (_level < _depth)
                    {
                        _level++;
                        WriteObject(prefix, item);
                        _level--;
                    }
                }
                else
                {
                    WriteObject(prefix, item);
                }
            }
        }

        /// <summary>
        /// Escribe un valor
        /// </summary>
        /// <param name="o"></param>
        private void WriteValue(object o)
        {
            if (o == null)
            {
                Write("null");
            }
            else if (o is DateTime)
            {
                Write(((DateTime)o).ToShortDateString());
            }
            else if (o is ValueType || o is string)
            {
                Write(o.ToString());
            }
            else if (o is IEnumerable)
            {
                Write("...");
            }
            else
            {
                Write("{ }");
            }
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
