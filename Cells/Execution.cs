using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq;

namespace EpplusTemplatesCS
{
    class Execution
    {
        public static Dictionary<string, string> GetMethodNames()
        {
            //Buscar sólo los métodos de ejecución de las configuraciones, que son públicas, estáticas y que no regresan valor (public static void)
            MethodInfo[] method_names = typeof(Execution).GetMethods(BindingFlags.Public | BindingFlags.Static).Where<MethodInfo>((method) => method.ReturnType.FullName == "System.Void").ToArray();
            Dictionary<string, string> methods = new Dictionary<string, string>();
            foreach (MethodInfo m in method_names)
            {
                string key = m.Name.Replace("_", "-"); //nombre de la configuración para su búsqueda
                string value = m.Name; //nombre de la función (con el guión bajo) y de obtención inmediata

                methods.Add(key, m.Name);
            }

            return methods;
        }
        public static List<PropertyNotNull> GetPropertiesNotNull(StyleConfig config)
        {
            List<PropertyNotNull> propsNotNull = new List<PropertyNotNull>();
            foreach (PropertyInfo prop in config.GetType().GetProperties())
            {
                bool propIsNotNull = prop.GetValue(config) != null;
                if (propIsNotNull)
                {
                    string key = prop.Name;
                    object value = prop.GetValue(config);

                    PropertyNotNull propertyNotNull = new PropertyNotNull() { name = key.Replace("_", "-"), value = value };
                    propsNotNull.Add(propertyNotNull);
                }
            }
            return propsNotNull;
        }

        #region CONFIGURATION_KEYS_FUNCTIONS
        public static void background_color(ExcelWorksheet Worksheet, string value, string range)
        {
            Worksheet.Cells[range].Style.Fill.SetBackground(ToolKits.Colors.ConvertToColor(value));
        }
        public static void bold(ExcelWorksheet Worksheet, bool value, string range)
        {
            Worksheet.Cells[range].Style.Font.Bold = value;
        }
        public static void border_color(ExcelWorksheet Worksheet, string value, string range)
        {
            //obtenemos la cantidad de colores que se metan separados por un espacio
            //MatchCollection parameters = Regex.Matches(value, @"\S+");
            MatchCollection parameters = Regex.Matches(value, @"([^\s][argb(].*[)])|(#?\w+)");
            Color[] colors = new Color[parameters.Count];

            if (parameters.Count > 0)
            {
                for (int i = 0; i < parameters.Count; i++)
                {
                    //colors[i] = ToolKits.Colors.ConvertToColor(FindCoincidence(parameters[i].Value, coincidences));
                    colors[i] = ToolKits.Colors.ConvertToColor(parameters[i].Value);
                }

            }
            else { return; }

            Border border = Worksheet.Cells[range].Style.Border;
            ExcelBorderItem top = border.Top;
            ExcelBorderItem right = border.Right;
            ExcelBorderItem bottom = border.Bottom;
            ExcelBorderItem left = border.Left;

            //asignar estilo de borde si no tiene
            if (
                border.Top.Style == ExcelBorderStyle.None &&
                border.Right.Style == ExcelBorderStyle.None &&
                border.Bottom.Style == ExcelBorderStyle.None &&
                border.Left.Style == ExcelBorderStyle.None &&
                colors.Length == 1
                )
            {
                border.BorderAround(Dictionaries.GetBorderStyle("default"), colors[0]);
                return;
            }
            else
            {
                ExcelBorderStyle _default = Dictionaries.GetBorderStyle("default");   //default

                //asignar estilo por default a algún borde que no tenga
                if (top.Style == ExcelBorderStyle.None) top.Style = _default;
                if (right.Style == ExcelBorderStyle.None) right.Style = _default;
                if (bottom.Style == ExcelBorderStyle.None) bottom.Style = _default;
                if (left.Style == ExcelBorderStyle.None) left.Style = _default;
            }


            if (colors.Length == 1)
            {
                top.Color.SetColor(colors[0]);
                right.Color.SetColor(colors[0]);
                bottom.Color.SetColor(colors[0]);
                left.Color.SetColor(colors[0]);
                return;
            }
            else if (colors.Length == 2)
            {
                top.Color.SetColor(colors[0]);
                right.Color.SetColor(colors[1]);
                bottom.Color.SetColor(colors[0]);
                left.Color.SetColor(colors[1]);
                return;
            }
            else if (colors.Length == 3)
            {
                top.Color.SetColor(colors[0]);
                right.Color.SetColor(colors[1]);
                bottom.Color.SetColor(colors[2]);
                left.Color.SetColor(colors[1]);
                return;
            }
            else if (colors.Length == 4)
            {
                top.Color.SetColor(colors[0]);
                right.Color.SetColor(colors[1]);
                bottom.Color.SetColor(colors[2]);
                left.Color.SetColor(colors[3]);
                return;
            }

        }
        public static void border_top_color(ExcelWorksheet Worksheet, string value, string range)
        {
            ExcelBorderItem border = Worksheet.Cells[range].Style.Border.Top;
            if (border.Style == ExcelBorderStyle.None)
                border.Style = Dictionaries.GetBorderStyle("default");

            border.Color.SetColor(ToolKits.Colors.ConvertToColor(value));
        }
        public static void border_right_color(ExcelWorksheet Worksheet, string value, string range)
        {
            ExcelBorderItem border = Worksheet.Cells[range].Style.Border.Right;
            if (border.Style == ExcelBorderStyle.None)
                border.Style = Dictionaries.GetBorderStyle("default");

            border.Color.SetColor(ToolKits.Colors.ConvertToColor(value));
        }
        public static void border_bottom_color(ExcelWorksheet Worksheet, string value, string range)
        {
            ExcelBorderItem border = Worksheet.Cells[range].Style.Border.Bottom;
            if (border.Style == ExcelBorderStyle.None)
                border.Style = Dictionaries.GetBorderStyle("default");

            border.Color.SetColor(ToolKits.Colors.ConvertToColor(value));
        }
        public static void border_left_color(ExcelWorksheet Worksheet, string value, string range)
        {
            ExcelBorderItem border = Worksheet.Cells[range].Style.Border.Left;
            if (border.Style == ExcelBorderStyle.None)
                border.Style = Dictionaries.GetBorderStyle("default");

            border.Color.SetColor(ToolKits.Colors.ConvertToColor(value));
        }
        public static void border_style(ExcelWorksheet Worksheet, string value, string range)
        {
            //obtenemos la cantidad de estilos que se metan separados por un espacio
            MatchCollection parameters = Regex.Matches(value, @"\S+");
            ExcelBorderStyle[] styles = new ExcelBorderStyle[parameters.Count];

            if (parameters.Count > 0)
            {
                for (int i = 0; i < parameters.Count; i++)
                {
                    //string _value = FindCoincidence(parameters[i].Value, coincidences);
                    //if (Dictionaries.border_style.ContainsKey(_value))
                    //{
                    //    styles[i] = Dictionaries.border_style[_value];
                    //}
                    styles[i] = Dictionaries.GetBorderStyle(parameters[i].Value);
                }
            }
            else { return; }

            Border border = Worksheet.Cells[range].Style.Border;
            ExcelBorderItem top = border.Top;
            ExcelBorderItem right = border.Right;
            ExcelBorderItem bottom = border.Bottom;
            ExcelBorderItem left = border.Left;

            if (styles.Length == 1)
            {
                top.Style = styles[0];
                right.Style = styles[0];
                bottom.Style = styles[0];
                left.Style = styles[0];
                return;
            }
            else if (styles.Length == 2)
            {
                top.Style = styles[0];
                right.Style = styles[1];
                bottom.Style = styles[0];
                left.Style = styles[1];
                return;
            }
            else if (styles.Length == 3)
            {
                top.Style = styles[0];
                right.Style = styles[1];
                bottom.Style = styles[2];
                left.Style = styles[1];
                return;
            }
            else
            {
                top.Style = styles[0];
                right.Style = styles[1];
                bottom.Style = styles[2];
                left.Style = styles[3];
                return;
            }
        }
        public static void border_top_style(ExcelWorksheet Worksheet, string value, string range)
        {
            //Worksheet.Cells[range].Style.Border.Top.Style = Dictionaries.border_style[value];
            Worksheet.Cells[range].Style.Border.Top.Style = Dictionaries.GetBorderStyle(value);
        }
        public static void border_right_style(ExcelWorksheet Worksheet, string value, string range)
        {
            //Worksheet.Cells[range].Style.Border.Right.Style = Dictionaries.border_style[value];
            Worksheet.Cells[range].Style.Border.Right.Style = Dictionaries.GetBorderStyle(value);
        }
        public static void border_bottom_style(ExcelWorksheet Worksheet, string value, string range)
        {
            //Worksheet.Cells[range].Style.Border.Bottom.Style = Dictionaries.border_style[value];
            Worksheet.Cells[range].Style.Border.Bottom.Style = Dictionaries.GetBorderStyle(value);
        }
        public static void border_left_style(ExcelWorksheet Worksheet, string value, string range)
        {
            //Worksheet.Cells[range].Style.Border.Left.Style = Dictionaries.border_style[value];
            Worksheet.Cells[range].Style.Border.Left.Style = Dictionaries.GetBorderStyle(value);
        }
        public static void color(ExcelWorksheet Worksheet, string value, string range)
        {
            Worksheet.Cells[range].Style.Font.Color.SetColor(ToolKits.Colors.ConvertToColor(value));
        }
        public static void font_family(ExcelWorksheet Worksheet, string value, string range)
        {
            Worksheet.Cells[range].Style.Font.Name = value;
        }
        public static void font_size(ExcelWorksheet Worksheet, float value, string range)
        {
            if (value > 0)
            {
                Worksheet.Cells[range].Style.Font.Size = value;
            }
        }
        public static void font_style(ExcelWorksheet Worksheet, string value, string range)
        {
            if (value.ToLower().Equals("italic") | value.ToLower().Equals("oblique")) italic(Worksheet, true, range);
        }
        public static void italic(ExcelWorksheet Worksheet, bool value, string range)
        {
            Worksheet.Cells[range].Style.Font.Italic = value;
        }
        public static void number_format(ExcelWorksheet Worksheet, string value, string range)
        {
            Worksheet.Cells[range].Style.Numberformat.Format = value;
        }
        public static void text_align(ExcelWorksheet Worksheet, string value, string range)
        {
            //Worksheet.Cells[range].Style.HorizontalAlignment = Dictionaries.text_align[value];
            Worksheet.Cells[range].Style.HorizontalAlignment = Dictionaries.GetTextAlign(value);
        }
        public static void text_wrap(ExcelWorksheet Worksheet, object value, string range)
        {
            bool boolean_value = false;
            if (Boolean.TryParse(value.ToString(), out boolean_value))
            {
                Worksheet.Cells[range].Style.WrapText = boolean_value;
                return;
            }

            Worksheet.Cells[range].Style.WrapText = value.ToString().ToLower().Equals("wrap"); //wrap = true, nowrap = false
        }
        public static void vertical_align(ExcelWorksheet Worksheet, string value, string range)
        {
            if (Dictionaries.vertical_align.ContainsKey(value))
                Worksheet.Cells[range].Style.VerticalAlignment = Dictionaries.vertical_align[value];
        }
        public static void width(ExcelWorksheet Worksheet, double value, string range)
        {
            string _range = Regex.Replace(range, "[0-9]", "");
            string[] columns = _range.Split(":");
            foreach (string column in columns)
            {
                int _column = ToolKits.General.N(column);
                Worksheet.Column(_column).Width = value;
            }
        }
        public static void height(ExcelWorksheet Worksheet, double value, string range)
        {
            string _range = Regex.Replace(range, "[A-Za-z]", "");
            string[] rows = _range.Split(":");
            foreach (string row in rows)
            {
                int _row = int.Parse(row);
                Worksheet.Row(_row).Height = value;
            }
        }
        #endregion CONFIGURATION_KEYS_FUNCTIONS

        #region OTHER_CONFIGURATION_KEYS_FUNCTIONS
        public static void merge(ExcelWorksheet Worksheet, bool value, string range)
        {
            Worksheet.Cells[range].Merge = value;
        }
        #endregion OTHER_CONFIGURATION_KEYS_FUNCTIONS
    }
}
