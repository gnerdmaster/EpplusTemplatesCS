using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace EpplusTemplatesCS
{
    public class Tables
    {
        #region Propiedades
        public ExcelWorksheet Worksheet { get; set; }
        #endregion

        public void Table(DataTable values, int column_init, int row_init, StyleConfig configurations = null)
        {
            //saber si el cell_init es correcto
            //bool is_correct_cell_init = Regex.IsMatch(cell_init, "^[a-zA-Z]+[0-9]+$");
            //if (!is_correct_cell_init) return;
            Worksheet.Cells[row_init, column_init].LoadFromDataTable(values, true);

        }
    }
    public class TableValue
    {
        public List<object[]> Headers { get; set; }
        public List<object[]> Body { get; set; }
        public List<object[]> Footers { get; set; }
    }
}
