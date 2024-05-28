using System;
using System.Collections.Generic;
using System.Text;

namespace EpplusTemplatesCS
{
    /// <summary>
    /// Clase para gestionar las posiciones de celda de algún objecto CellPosition
    /// </summary>
    public class CellPosition
    {
        private int _column_init = 1;
        private int _row_init = 1;
        private int _column_finish = 1;
        private int _row_finish = 1;
        /// <summary>
        /// Columna de inicio
        /// </summary>
        public int ColumnInit
        {
            get
            {
                return _column_init;
            }
            set
            {
                _column_init = value <= 0 ? _column_init : value;
            }
        }
        /// <summary>
        /// Fila de inicio
        /// </summary>
        public int RowInit
        {
            get
            {
                return _row_init;
            }
            set
            {
                _row_init = value <= 0 ? _row_init : value;
            }
        }
        /// <summary>
        /// Columna final
        /// </summary>
        public int ColumnFinish
        {
            get
            {
                return _column_finish <= _column_init ? _column_init : _column_finish;
            }
            set
            {
                _column_finish = value <= 0 ? _column_finish : value;
            }
        }
        /// <summary>
        /// Fila final
        /// </summary>
        public int RowFinish
        {
            get
            {
                return _row_finish <= _row_init ? _row_init : _row_finish;
            }
            set
            {
                _row_finish = value <= 0 ? _row_finish : value;
            }
        }

        #region METODOS
        #endregion METODOS
    }
}
