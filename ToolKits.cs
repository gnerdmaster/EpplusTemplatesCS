using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;

namespace EpplusTemplatesCS
{
    /// <summary>
    /// EpplusTemplatesCS - Clase Caja de Herramientas
    /// </summary>
    public static class ToolKits
    {
        #region Extensions
        /// <summary>
        /// Remplaza texto de la cadena por posiciones iniciando por el #0. 
        /// <br/>Ej. 
        /// <br/>string cadena = "#0 es un #1 de #2";
        /// <br/>string nueva_cadena = cadena._KeyReplace("esto", "ejemplo", "reemplazo", ...);
        /// <br/>//nueva cadena => "esto es un ejemplo de reemplazo"
        /// </summary>
        /// <param name="key">Template</param>
        /// <param name="replacements">Reemplazos</param>
        /// <returns></returns>
        public static string _KeyReplace(this string key, params string[] replacements)
        {
            //reemplaza por posiciones Ej. 0,1,2,3...
            for (int i = 0; i < replacements.Length; i++)
            {
                key = key.Replace("#" + i.ToString(), replacements[i]);
            }

            return key; //nuevo string a utilizar
        }

        /// <summary>
        /// Retorna un nuevo elemento CellPosition con las posiciones máximas entre dos elementos CellPosition
        /// </summary>
        /// <param name="pos_target">Posición destino</param>
        /// <param name="pos_source">Posición de origen</param>
        /// <returns>new CellPosition()</returns>
        public static CellPosition Enlarge(this CellPosition pos_target, CellPosition pos_source)
        {
            if (pos_source == null) return null;
            if (pos_target == null) return pos_target;

            pos_target.ColumnFinish = pos_source.ColumnFinish > pos_target.ColumnFinish ? pos_source.ColumnFinish : pos_target.ColumnFinish;
            pos_target.RowFinish = pos_source.RowFinish > pos_target.RowFinish ? pos_source.RowFinish : pos_target.RowFinish;
            return pos_target;
        }

        /// <summary>
        /// Obtiene el valor con un key sin tener qué verificar si existe o no.
        /// <br/>Si no hay valor, entonces manda el valor predeterminado del diccionario.
        /// </summary>
        /// <typeparam name="K">Key</typeparam>
        /// <typeparam name="V">Value</typeparam>
        /// <param name="dictionary">Diccionario</param>
        /// <param name="key">Key de entrada</param>
        /// <returns></returns>
        public static V _GetValue<K, V>(this IDictionary<K, V> dictionary, K key)
        {
            if (dictionary.ContainsKey(key))
            {
                return dictionary[key];
            }

            return default(V);
        }
        /// <summary>
        /// Obtiene el valor con un key sin tener qué verificar si existe o no.
        /// <br/>Si no hay valor, entonces manda el valor predeterminado del diccionario.
        /// </summary>
        /// <typeparam name="K">Key</typeparam>
        /// <typeparam name="V">Value</typeparam>
        /// <param name="dictionary">Diccionario</param>
        /// <param name="key">Key de entrada</param>
        /// <param name="_default">Valor por defecto</param>
        /// <returns></returns>
        public static V _GetValue<K, V>(this IDictionary<K, V> dictionary, K key, V _default)
        {
            if (dictionary.ContainsKey(key))
            {
                return dictionary[key];
            }

            return _default;
        }
        #endregion Extensions
        public static class Text
        {
            public static string RenameColor(string color_name)
            {
                return Separator(color_name, "-", "[A-Z]", true).ToLower();
            }
            public static string Separator(string text, string separator = "", string pattern = "", bool ignore_first_char = false)
            {
                int size_separator = separator.Length;

                bool haveCapitalLetter = Regex.IsMatch(ignore_first_char ? text.Substring(1) : text, pattern);   //no tomará en cuenta la primera letra para que busque la segunda "BlueViolet" => "lueViolet" => "V" => true
                if (haveCapitalLetter)
                {
                    for (int i = 0; i < text.Length; i++)
                    {
                        if (char.IsUpper(text[i]) && i > 0)
                        {
                            text = text.Insert(i, separator);
                            if (!string.IsNullOrEmpty(separator)) i += size_separator; //si no es nulo o vacío entonces se suma posiciones (conforme al tamaño del separador) porque se le ha agregado uno o más caracteres de separador en el mismo texto
                        }
                    }
                }

                return text;
            }
        }
        public static class Colors
        {
            /// <summary>
            /// Obtiene los 3 o 4 valores ARGB.
            /// </summary>
            /// <param name="argb"></param>
            /// <returns></returns>
            public static Color StringARGBToColor(string argb)
            {
                MatchCollection rgb = Regex.Matches(argb, "[0-9]+");

                int param_count = rgb.Count;
                Color color_argb = new Color();
                if (param_count == 3)
                {
                    //sólo tiene rgb (1,2,3)
                    color_argb = Color.FromArgb(Int32.Parse(rgb[0].Value), Int32.Parse(rgb[1].Value), Int32.Parse(rgb[2].Value));
                }
                else if (param_count == 4)
                {
                    //tiene el parámetro alfa (1,2,3,4)
                    color_argb = Color.FromArgb(Int32.Parse(rgb[0].Value), Int32.Parse(rgb[1].Value), Int32.Parse(rgb[2].Value), Int32.Parse(rgb[3].Value));
                }

                return color_argb;
            }

            /// <summary>
            /// Devuelve un objeto Color a partir del color ingresado, ya sea por hexadecimal, (a)rgb o nombre. Ej.
            /// <br/> <c>yellow | #ffffff | rgb(255,255,255) | argb(255,255,255)</c>
            /// </summary>
            /// <param name="color">Hexadecimal <b>#ffffff</b> | (A)RGB <b>rgb(255,255,255)</b> | Nombre <b>yellow</b></param>
            /// <returns>System.Drawing.Color</returns>
            public static Color ConvertToColor(string color)
            {
                string patternARGB = @"^((a|A)?(r|R)(g|G)(b|B))?\(?([01]?\d\d?|2[0-4]\d|25[0-5])(\W+)([01]?\d\d?|2[0-4]\d|25[0-5])\W+(([01]?\d\d?|2[0-4]\d|25[0-5])\)?)$";

                //nombre de color, si no existe => white predeterminado
                Color _color = Dictionaries.GetColor(color);

                if (color.StartsWith("#"))
                {
                    //códigos HEXADECIMALES de colores HTML
                    _color = ColorTranslator.FromHtml(color);
                }
                else if (Regex.IsMatch(color.Replace(" ", ""), patternARGB))
                {
                    //códigos ARGB y RGB
                    _color = ToolKits.Colors.StringARGBToColor(color.Replace(" ", ""));
                }

                return _color;
            }
        }

        public static class General
        {
            /// <summary>
            /// Convierte número (posición de columna) a letra
            /// </summary>
            /// <param name="column_number"></param>
            /// <returns></returns>
            public static string L(int column_number)
            {
                int a = 0;
                int b = 0;
                string letter = "";
                do
                {
                    a = (column_number - 1) / 26;
                    b = (column_number - 1) % 26;
                    letter = (char)(b + 65) + letter;
                    column_number = a;
                } while (column_number > 0);

                return letter;
            }
            /// <summary>
            /// Convierte letra (columna) a entero (posición de columna)
            /// </summary>
            /// <param name="column_letter"></param>
            /// <returns></returns>
            public static int N(string column_letter)
            {
                //string column_letter = "acz";
                string abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

                int column_number = 0;
                foreach (char letter in column_letter.ToUpper())
                {
                    column_number = (abc.IndexOf(letter) != -1 ? abc.IndexOf(letter) + 1 : 0) + column_number * abc.Length;
                }
                return column_number;
            }
        }
    }
}
