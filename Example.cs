using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EpplusTemplatesCS
{
    class Example
    {
        public static void EjemploCellPosition(Cells C)
        {
            CellPosition test = new CellPosition() { ColumnInit = C.N("a") };
            C.Cell("hola", test);

            Dictionary<string, CellPosition> positions = new Dictionary<string, CellPosition>();
            positions.Add("test", test);

            //positions["test"].ColumnInit = 1;



            int total_new_columns = 2;
            int position_new_columns = 1; //aumentando 2 columnas nuevas en la posición 1
            C.Worksheet.InsertColumn(position_new_columns, total_new_columns);


            positions.AsParallel().ForAll(x =>
            {
                if (x.Value.ColumnInit >= position_new_columns)
                {
                    x.Value.ColumnInit += total_new_columns;
                }
                else
                {
                    if (x.Value.ColumnFinish >= position_new_columns)
                    {
                        x.Value.ColumnFinish += total_new_columns;
                    }
                }
            });

            //foreach (CellPosition position in positions.Values)
            //{
            //    if (position.ColumnInit >= new_column_in_position)
            //    {
            //        position.ColumnInit += total_new_columns;
            //    }
            //    else
            //    {
            //        if (position.ColumnFinish >= new_column_in_position)
            //        {
            //            position.ColumnFinish += total_new_columns;
            //        }
            //    }
            //}

        }
    }
}
