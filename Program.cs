using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace EpplusTemplatesCS
{
    class Program
    {
        static void Main(string[] args)
        {
            Export();
        }
        public static void Export()
        {
            try
            {
                //matamos cualquier proceso de microsoft excel que exista
                var p = new Process();
                foreach (Process proc in Process.GetProcessesByName("EXCEL"))
                {
                    string nombreExcel = proc.MainWindowTitle.Trim();
                    if (nombreExcel.Contains("test.xlsx"))
                    {
                        proc.Kill();
                    }
                }

                var fileInfo = new FileInfo(@".\test.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //para que no salga la excepción de la comercial
                using (var package = new ExcelPackage())
                {
                    ExcelWorksheet excelWorksheet = package.Workbook.Worksheets.Add("NEW TESTING");
                    PaintCells(excelWorksheet);
                    package.File = fileInfo;
                    package.Save();

                    p.StartInfo = new ProcessStartInfo(@".\test.xlsx")
                    {
                        UseShellExecute = true
                    };
                    p.Start();

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error" + e.Message);
            }
        }

        private static void PaintCells(ExcelWorksheet excelWorksheet)
        {
            //MODO DE USO:
            //1.- Invocar al creador de Celdas (clase) y asignando el worksheet en donde modificará
            Cells C = new Cells()
            {
                Worksheet = excelWorksheet
            };
            bool execution = false;
            if (execution)
            {
                //2.- Llamar el método Cell e insertar los datos requeridos
                C.Cell("C.Cell(\"valor\", \"A1\");", "A1");
                C.Cell(22, "A2"); //aceptación de otros valores que no sean cadena de texto
                C.Cell("=1+3", "A3"); //función
                C.Cell("Celdas sin combinar, se repite el valor en el rango de celda establecido A5:B6", "A5:B6");
                C.Cell("Celdas combinadas", "+a7:b8");

                C.Cell("text-align: start", "a9", "text-align: start");
                C.Cell("text-align: center", "a10", "text-align: center");
                C.Cell("text-align: center-continuous", "a11", "text-align: center-continuous");
                C.Cell("text-align: end", "a12", "text-align: end");
                C.Cell("text-align: justify", "a13", "text-align: justify");
                C.Cell("text-align: fill", "a14", "text-align: fill");
                C.Cell("text-align: distributed", "a15", "text-align: distributed");
                C.Cell("text-align: general", "a16", "text-align: general");

                C.Cell("vertical-align: top", "a17", "vertical-align: top");
                C.Cell("vertical-align: middle", "a18", "vertical-align: middle");
                C.Cell("vertical-align: bottom", "a19", "vertical-align: bottom");
                C.Cell("vertical-align: distributed", "a20", "vertical-align: distributed");
                C.Cell("vertical-align: justify", "a21", "vertical-align: justify");

                C.Cell("background-color: blue", "a22", "background-color: blue");
                C.Cell("background-color: brown", "a23", "background-color: brown");
                C.Cell("background-color: lime", "a24", "background-color: lime");
                C.Cell("background-color: yellow", "a25", "background-color: yellow");
                C.Cell("background-color: yellow-green", "a26", "background-color: yellow-green");
                C.Cell("background-color: RGB( 0, 255, 0)", "a27", "background-color: RGB( 0, 255, 0)");

                C.Cell("BACKGROUND-COLOR:YELLOW", "A28", "BACKGROUND-COLOR:YELLOW");
                C.Cell("BACKGROUND-COLOR:Yellow", "A29", "BACKGROUND-COLOR:Yellow");
                C.Cell("BACKGROUND-COLOR:YellowGreen", "A30", "BACKGROUND-COLOR:YellowGreen");
                C.Cell("BACKGROUND-COLOR:yellow-green", "A31", "BACKGROUND-COLOR:yellow-green");
                C.Cell("BACKGROUND-COLOR:yellow", "A32", "BACKGROUND-COLOR:yellow");

                //borders
                C.Cell("border-style: medium", "b26", "border-style: medium");

                C.Cell("border-top-color: blue", "b28", "border-top-color: blue");
                C.Cell("border-right-color: blue", "b29", "border-right-color: blue");
                C.Cell("border-bottom-color: blue", "b30", "border-bottom-color: blue");
                C.Cell("border-left-color: blue", "b31", "border-left-color: blue");

                C.Cell("border-color: #BAA7A3", "B32", "border-color: #BAA7A3");
                C.Cell("border-color: #BAA7A3 yellow", "B34", "border-color: #BAA7A3 yellow");
                C.Cell("border-color: #BAA7A3 yellow blue", "B36", "border-color: #BAA7A3 yellow blue");
                C.Cell("border-color: #BAA7A3 yellow blue rgb(0,255,0)", "B38", "border-color: #BAA7A3 yellow blue rgb(0,255,0)");

                C.Cell("border-style: double", "B40", "border-style: double");
                C.Cell("border-style: double dotted", "B42", "border-style: double dotted");
                C.Cell("border-style: double dotted dashed", "B44", "border-style: double dotted dashed");
                C.Cell("border-style: double dotted dashed thick", "B46", "border-style: double medium dotted dashed");

                C.Cell("border-top-style: double", "B48", "border-top-style: double");
                C.Cell(null, "B48", "border-top-style: double");

                //==== CREAR CELDAS CON UN OBJETO Configuraciones(Documentación incluída con más entendimiento y control del código)
                C.Cell("new StyleConfig() { text_align = \"center\", vertical_align = \"middle\", border_color = \"yellow red #FF0033 rgb(0,255,255)\" }", "D2", new StyleConfig() { text_align = "center", vertical_align = "middle", border_color = "yellow red #FF0033 rgb(0,255,255)" });
                C.Cell("new StyleConfig text with style 2", "D4", new StyleConfig() { border_style = "single" });
                C.Cell("new StyleConfig text with style 3", "D6", new StyleConfig() { border_style = "single thick" });
                C.Cell("new StyleConfig text with style 4", "D8", new StyleConfig() { border_style = "single thick medium-dash-dot" });
                C.Cell("new StyleConfig text with style 4", "D10", new StyleConfig() { border_style = "single thick medium-dash-dot dot" });


                C.Cell("new StyleConfig(){ bold = true}", "D12", new StyleConfig() { bold = true });
                C.Cell("BOLD: TRUE", "D13", "BOLD: TRUE");

                C.Cell(10, "D15", new StyleConfig() { number_format = "$###,###,##0.00" });

                C.Cell("text-wrap: wrap, ejemplo de un text-wrap", "D16", "text-wrap: wrap");
                C.Cell("text-wrap: nowrap, ejemplo de un text-wrap", "D17", "text-wrap: nowrap");
                C.Cell("new StyleConfig() { text_wrap = true }, ejemplo de un text-wrap", "D18", new StyleConfig() { text_wrap = true });
                C.Cell("new StyleConfig() { text_wrap = false }, ejemplo de un text-wrap", "D19", new StyleConfig() { text_wrap = false });

                C.Cell("new StyleConfig() { font_size = 20 }", "D21", new StyleConfig() { font_size = 20 });
                C.Cell("new StyleConfig() { font_family = \"Arial\" }", "D22", new StyleConfig() { font_family = "Arial" });
                C.Cell("font-family: Courier New", "D23", "font-family: Courier New");
                C.Cell("font-family: Elephant", "D24", "font-family: Elephant");
                C.Cell("font-style: italic", "D25", "font-style: italic");
                C.Cell("new StyleConfig() { font_style = \"italic\" }", "D26", new StyleConfig() { font_style = "italic" });
                C.Cell("color: orange", "D27", "color: orange");
                C.Cell("new StyleConfig() { color = \"#00ff55\" }", "D28", new StyleConfig() { color = "#00ff55" });
                C.Cell("color: rgb(0,255,0)", "D29", "color: rgb(0,255,0)");

                for (int i = 1; i < 27; i++)
                {
                    excelWorksheet.Column(i).AutoFit();
                }
            }
            else
            {
                //sólo testing
                C.Cell(null, "A1:B2", new StyleConfig() { height = 50 });
                C.Cell(null, "A1:B2", new StyleConfig() { width = 30 });
            }
        }

        private static void CreateTable(Cells C)
        {
            //Posiciones de Celdas de cada uno de los componentes
            CellPosition _table = new CellPosition()
            {
                RowInit = 1,
                ColumnInit = ToolKits.General.N("E")
            };
            CellPosition _rowheaders = new CellPosition()
            {
                RowInit = _table.RowInit + 1,
                ColumnInit = _table.ColumnInit,
                ColumnFinish = _table.ColumnInit
            };
            CellPosition _headers = new CellPosition()
            {
                RowInit = _table.RowInit,
                ColumnInit = _rowheaders.ColumnFinish + 1,
                RowFinish = _table.RowInit
            };
            CellPosition _subheaders = new CellPosition()
            {
                RowInit = _headers.RowFinish + 1,
                ColumnInit = _headers.ColumnInit,
                RowFinish = _headers.RowFinish + 1
            };
            CellPosition _data = new CellPosition()
            {
                RowInit = _rowheaders.RowInit,
                ColumnInit = _headers.ColumnInit
            };
            CellPosition _subtable = new CellPosition()
            {
                RowInit = _subheaders.RowInit + 1,
                ColumnInit = _rowheaders.ColumnInit + 1
            };


            List<string> RowHeaders = new List<string>() {
                "ROW_HEADER_01",
                //"ROW_HEADER_02"
            };

            C.Cell("ESTATUS", _rowheaders.ColumnInit, _rowheaders.RowInit); //título de los estatus
            for (int index_rowheader = 0; index_rowheader < RowHeaders.Count; index_rowheader++)
            {
                List<string> Headers = new List<string>() {
                    "HEADER 1",
                    "HEADER 2",
                    "HEADER 3",
                };
                List<string> SubHeaders = null;
                List<string> SubtableTitles = null;
                List<object[]> DataSubtable = null;

                for (int index_header = 0; index_header < Headers.Count; index_header++)
                {
                    SubHeaders = new List<string>
                    {
                        "H1 - SUBTITLE 1",
                        "H1 - SUBTITLE 2",
                        //"H1 - SUBTITLE 3"
                    };

                    SubtableTitles = new List<string>()
                    {
                        "TITLE_SUBTABLE_01", "TITLE_SUBTABLE_02", "TITLE_SUBTABLE_02"
                    };

                    DataSubtable = new List<object[]>() {
                        new object[]{"NAME_01", "DESCRIPTION_01_01", "DESCRIPTION_02_01"},
                        new object[]{"NAME_02", "DESCRIPTION_01_02", "DESCRIPTION_02_02"},
                        new object[]{"NAME_03", "DESCRIPTION_01_03", "DESCRIPTION_02_03"},
                    };

                    //Pre-cálculos
                    //cantidad de títulos del subheader
                    int total_titles_subheader = SubHeaders.Count;
                    int total_titles_subtable = SubtableTitles.Count;

                    //si hay la misma cantidad de subheaders que los títulos de la subtable => cada título de subtable cabe en la misma columna que un titulo de subheader
                    bool is_equal = total_titles_subheader == total_titles_subtable;
                    if (is_equal)
                    {
                        total_titles_subheader = 1;
                        total_titles_subtable = 1;
                    }


                    //=========== CREACIÓN DE COMPONENTES ==============

                    /*1.- Imprimir la subtabla*/
                    CreateSubtable(C, ref _subtable, SubtableTitles, DataSubtable, total_titles_subheader);

                    /*2.- Imprimir las cabeceras y subcabeceras, si existen*/
                    CreateHeaders(C, ref _headers, ref _subheaders, Headers, index_header, SubHeaders, total_titles_subheader, total_titles_subtable);

                    /*3.- Asignación y reasignación de coordenadas*/

                    //registrando la columna final de todos los componentes, es decir de la tabla completa
                    _table.ColumnFinish = _headers.ColumnFinish;
                    _table.RowFinish = _subtable.RowFinish;

                    //reiniciando las columnas de inicio de todos los componentes que se deben de ciclar
                    _headers.ColumnInit = _table.ColumnFinish + 1;
                    _subheaders.ColumnInit = _headers.ColumnInit;
                    _subtable.ColumnInit = _subheaders.ColumnInit;
                }

                //4.- Imprimir la cabecera lateral Estatus (con el pre-cálculo de las coordenadas de la altura máxima en cantidad de filas de la subtabla de las ots, que sería el valor de _table.RowFinish + 1)


                //Crea Cabeceras lateral izquierda
                _rowheaders.RowFinish = _rowheaders.RowInit + index_rowheader + 1;
                C.Cell(RowHeaders[index_rowheader], _rowheaders.ColumnInit, _rowheaders.RowFinish, _rowheaders.ColumnInit, _table.RowFinish, new StyleConfig() { merge = true });
            }
        }

        private static void CreateHeaders(Cells C, ref CellPosition _headers, ref CellPosition _subheaders, List<string> Headers, int index_header, List<string> SubHeaders, int total_titles_subheader, int total_titles_subtable)
        {
            //subheaders
            for (int index_subheader = 0; index_subheader < SubHeaders.Count; index_subheader++)
            {
                _subheaders.ColumnFinish = _subheaders.ColumnInit + index_subheader * total_titles_subtable;
                C.Cell(SubHeaders[index_subheader], _subheaders.ColumnFinish, _subheaders.RowInit, _subheaders.ColumnFinish + (total_titles_subtable - 1), _subheaders.RowInit, new StyleConfig() { merge = true });
            }

            //actualizar la columna final del subheaders
            _subheaders.ColumnFinish += (total_titles_subtable - 1);

            _headers.ColumnFinish = _subheaders.ColumnFinish;    //actualizar la columna final de los headers
            C.Cell(Headers[index_header], _headers.ColumnInit, _headers.RowInit, _headers.ColumnFinish, _headers.RowInit, new StyleConfig() { merge = true }); ;
        }

        private static void CreateSubtable(Cells C, ref CellPosition _subtable, List<string> subtables_titles, List<object[]> data_subtable, int total_titles_subheader)
        {
            //1.1.- Títulos
            for (int index_title = 0; index_title < subtables_titles.Count; index_title++)
            {
                _subtable.ColumnFinish = _subtable.ColumnInit + index_title * total_titles_subheader;
                C.Cell(subtables_titles[index_title], _subtable.ColumnFinish, _subtable.RowInit, _subtable.ColumnFinish + (total_titles_subheader - 1), _subtable.RowInit, new StyleConfig() { merge = true });
            }
            //1.2.- Datos
            for (int index_row_subtable = 0; index_row_subtable < data_subtable.Count; index_row_subtable++)
            {
                _subtable.RowFinish = _subtable.RowInit + index_row_subtable + 1;

                for (int index_column_subtable = 0; index_column_subtable < data_subtable[index_row_subtable].Length; index_column_subtable++)
                {
                    _subtable.ColumnFinish = _subtable.ColumnInit + index_column_subtable * total_titles_subheader;
                    C.Cell(data_subtable[index_row_subtable][index_column_subtable], _subtable.ColumnFinish, _subtable.RowFinish, _subtable.ColumnFinish + (total_titles_subheader - 1), _subtable.RowFinish, new StyleConfig() { merge = true });
                }

                //actualizar la columna Final del subtable
                _subtable.ColumnFinish += (total_titles_subheader - 1);
            }
        }
    }

    public class Header
    {
        public string Title { get; set; }
        public CellPosition Positions { get; set; }

    }
}
