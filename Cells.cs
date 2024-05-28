using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;


namespace EpplusTemplatesCS
{
    public class Cells
    {
        #region Propiedades
        public ExcelWorksheet Worksheet { get; set; }
        #endregion Propiedades

        #region Métodos Pricipales
        public void Cell(object value, string range, string style_config)
        {
            _Cell(value, range, style_config);
        }
        public void Cell(object value, string range, StyleConfig style_config = null)
        {
            _Cell(value, range, style_config);
        }
        public void Cell(object value, int column, int row, string style_config)
        {
            if (column <= 0) column++;
            if (row <= 0) row++;

            string range = $"{ToolKits.General.L(column)}{row}";
            _Cell(value, range, style_config);
        }
        public void Cell(object value, int column, int row, StyleConfig style_config = null)
        {
            if (column <= 0) column++;
            if (row <= 0) row++;

            string range = $"{ToolKits.General.L(column)}{row}";
            _Cell(value, range, style_config);
        }
        public void Cell(object value, int from_column, int from_row, int to_column, int to_row, string style_config)
        {
            if (from_column <= 0) from_column++;
            if (from_row <= 0) from_row++;
            if (to_column <= 0) to_column++;
            if (to_row <= 0) to_row++;

            string range = $"{ToolKits.General.L(from_column)}{from_row}:{ToolKits.General.L(to_column)}{to_row}";
            _Cell(value, range, style_config);
        }
        public void Cell(object value, int from_column, int from_row, int to_column, int to_row, StyleConfig style_config = null)
        {
            if (from_column <= 0) from_column++;
            if (from_row <= 0) from_row++;
            if (to_column <= 0) to_column++;
            if (to_row <= 0) to_row++;

            string range = $"{ToolKits.General.L(from_column)}{from_row}:{ToolKits.General.L(to_column)}{to_row}";
            _Cell(value, range, style_config);
        }
        public void Cell(object value, CellPosition position, string style_config)
        {
            if (position == null) return;

            string column_init = ToolKits.General.L(position.ColumnInit <= 0 ? 1 : position.ColumnInit);
            string column_finish = position.ColumnFinish <= 0 ? column_init : ToolKits.General.L(position.ColumnFinish);
            int row_init = position.RowInit <= 0 ? 1 : position.RowInit;
            int row_finish = position.RowFinish <= 0 ? row_init : position.RowFinish;

            _Cell(value, $"{column_init}{row_init}:{column_finish}{row_finish}", style_config);
        }
        public void Cell(object value, CellPosition position, StyleConfig style_config = null)
        {
            if (position == null) return;

            string column_init = ToolKits.General.L(position.ColumnInit <= 0 ? 1 : position.ColumnInit);
            string column_finish = position.ColumnFinish <= 0 ? column_init : ToolKits.General.L(position.ColumnFinish);
            int row_init = position.RowInit <= 0 ? 1 : position.RowInit;
            int row_finish = position.RowFinish <= 0 ? row_init : position.RowFinish;

            _Cell(value, $"{column_init}{row_init}:{column_finish}{row_finish}", style_config);
        }
        private void _Cell(object value, string range, object style_config = null)
        {
            bool isCombined = range.StartsWith("+");    //celdas convinadas o no
            range = isCombined ? range.Remove(0, 1) : range;    //quitar el signo de combinación del rango

            if (value != null)
            {
                //1. Asignar valor | función a la celda
                bool isFunction = value.GetType().FullName == "System.String" && (value.ToString().StartsWith("=") || value.ToString().StartsWith("+"));
                if (isFunction)
                {
                    Worksheet.Cells[range].Formula = value.ToString().Remove(0, 1); //quita el primer caracter que es el signo que determina si es una función o no (=, +, -)
                }
                else
                {
                    Worksheet.Cells[range].Value = value;
                }
            }

            //2. Manejar rango de celdas
            Worksheet.Cells[range].Merge = isCombined;

            //3.Aplicar configuraciones | estilos
            SetStyleConfig(range, style_config);
        }
        #endregion Metodos_Principales

        #region Metodos_secundarios
        private void SetStyleConfig(string range, object style_config)
        {
            if (style_config == null || string.IsNullOrEmpty(style_config.ToString())) return;  //si es nulo entonces salir del método

            //a.Obtener las configuraciones y setearlas en un objeto para su uso inmediato
            //1. Validar cada configuración que tenga valor(que no sea null) en un sólo tamo
            StyleConfig result_config = new StyleConfig();
            if (style_config.GetType().FullName == "System.String")
            {
                result_config = LexicalInterpreter(style_config.ToString());
            }
            else
            {
                result_config = (StyleConfig)style_config;
            }

            //Ejecución de las configuraciones
            ConfigurationExecution(range, result_config);
        }
        #endregion Metodos_secundarios

        #region LexicalInterpreter
        private StyleConfig LexicalInterpreter(string style_config)
        {
            if (string.IsNullOrEmpty(style_config)) return new StyleConfig();   //si no hay ninguna configuración entonces interrumpir la función

            //obtenemos las coincidencias
            string pattern = "([^:\\s]+)\\s*:\\s*[\"']?([^\"';]*)[\"']?";
            MatchCollection coincidences = Regex.Matches(style_config, pattern);

            StyleConfig config = new StyleConfig();

            //separamos key:value
            foreach (Match _coincidence in coincidences)
            {
                string key = _coincidence.Groups[1].Value;
                string value = _coincidence.Groups[2].Value;

                string name_property = key.Replace("-", "_").Trim().ToLower();

                //si un parámetro no existe => ignora el seteo
                PropertyInfo prop_info = config.GetType().GetProperty(name_property);

                if (prop_info != null)
                {
                    bool value_boolean = false;
                    int value_int = 0;
                    double value_double = 0.0;
                    float value_float = 0;

                    //Si los tipos de propiedades del valor entrante vs la clase del objeto no son iguales  => hacer comprobaciones de tipos
                    Type prop_type = prop_info.PropertyType;
                    Type value_type = value.GetType();
                    if (value_type != prop_type)
                    {
                        if (Boolean.TryParse(value, out value_boolean))
                        {
                            prop_info.SetValue(config, value_boolean);
                            continue;
                        }
                        else if (int.TryParse(value, out value_int))
                        {
                            prop_info.SetValue(config, value_int);
                            continue;
                        }
                        else if (Double.TryParse(value, out value_double))
                        {
                            prop_info.SetValue(config, value_double);
                            continue;
                        }
                        else if (float.TryParse(value, out value_float))
                        {
                            prop_info.SetValue(value, value_float);
                            continue;
                        }

                    }
                    prop_info.SetValue(config, value);
                }
            }


            return config;
        }
        #endregion LexicalInterpreter

        #region ConfigurationExecution
        private void ConfigurationExecution(string range, StyleConfig config)
        {
            //====== SUSTITUCIÓN DE CADENAS ======//
            //sólo propiedades de clase sin valor null
            List<PropertyNotNull> propsNotNull = Execution.GetPropertiesNotNull(config);

            //====== EXECUCIÓN DE CADA CONFIGURACIONES (limitadas sólo a las que se van a usar)======//
            foreach (PropertyNotNull property in propsNotNull)
            {
                object value = property.value;
                string name_function = property.name.Replace("-", "_");

                //ejecutar el método indicado para la configuración solicitada => si existe método => invocarlo
                MethodInfo function = typeof(Execution).GetMethod(name_function);
                if (function != null)
                    function.Invoke(null, new object[] { Worksheet, value, range });
                else continue;
            }
        }
        #endregion ConfigurationExecution

        #region Otros
        /// <summary>
        ///  Convierte número (posición de columna) a Letra
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public string L(int columnNumber)
        {
            return ToolKits.General.L(columnNumber);
        }

        /// <summary>
        /// Convierte letra (posición de columna) a Número
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public int N(string columnLetter)
        {
            return ToolKits.General.N(columnLetter);
        }
        #endregion Otros
    }

    
}
