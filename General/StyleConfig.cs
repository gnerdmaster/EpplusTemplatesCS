

namespace EpplusTemplatesCS
{
    /// <summary>
    /// Clase StyleConfig, tiene las propiedades necesarias para establecer los estilos y las configuraciones de la celda a imprimir.
    /// </summary>
    public class StyleConfig
    {
        /// <summary>
        /// ✔<c>Background-color</c> es un propiedad que define el color de fondo de un elemento, puede ser el valor de un color o la palabra clave <c>transparent</c>(<b>white</b> por defecto).
        /// <br/><b>Sintaxis</b>
        /// <br/><b><c>background-color: color | transparent</c></b>
        /// </summary>
        public string background_color { get; set; }
        /// <summary>
        /// La propiedad border permite definir en una única regla todos los bordes de los elementos seleccionados. Se puede utilizar border para definir el o los valores siguientes: <c><b>border-width</b>, <b>border-style</b>, <b>border-color</b>.</c>
        /// <br/><b>Sintaxis</b>
        /// <br/><c>border: [border-width || border-style || border-color] ;</c>
        /// </summary>
        public string border { get; set; }
        //public string border_top { get; set; }
        //public string border_right { get; set; }
        //public string border_bottom { get; set; }
        //public string border_left { get; set; }
        //public string border_width { get; set; }
        //public string border_top_width { get; set; }
        //public string border_right_width { get; set; }
        //public string border_bottom_width { get; set; }
        //public string border_left_width { get; set; }
        public string border_color { get; set; }
        /// <summary>
        /// ✔La propiedad <c>border-color</c> es un atajo para definir el color de los cuatro bordes de un elemento.
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>[ color || transparent ]{1,4}</c>
        /// </summary>
        public string border_top_color { get; set; }
        public string border_right_color { get; set; }
        public string border_bottom_color { get; set; }
        public string border_left_color { get; set; }
        public string border_style { get; set; }
        public string border_top_style { get; set; }
        public string border_right_style { get; set; }
        public string border_bottom_style { get; set; }
        public string border_left_style { get; set; }
        //public string border_inside_horizontal { get; set; }
        //public string border_inside_vertical { get; set; }
        //public string border_inside_horizontal_color { get; set; }
        //public string border_inside_vertical_color { get; set; }
        /// <summary>
        /// ✔La propiedad <c>font-size</c> especifica la dimensión de la letra.
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>font-size: lenght</c>
        /// </summary>
        public float? font_size { get; set; }
        /// <summary>
        /// ✔La propiedad <c>font-family</c> define una lista de fuentes o familias de fuentes.
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>font-family: [family | generic name]</c>
        /// </summary>
        public string font_family { get; set; }
        /// <summary>
        /// ✔La propiedad <c>font-style</c> permite definir el aspecto de una familia tipográfica entre los valores: normal, italic (cursiva)
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>font-family: [normal | italic]</c>
        /// </summary>
        public string font_style { get; set; }
        /// <summary>
        /// ✔La propiedad <c>color</c> selecciona el valor de color de primer plano del contenido de elemento de texto y decoraciones de texto.
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>color: [name | #ffffff | rgb(255,255,255)]</c>
        /// </summary>
        public string color { get; set; }
        /// <summary>
        /// ✔Fuente Cursiva
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>italic: [true | false]</c>
        /// </summary>
        public bool? italic { get; set; }
        /// <summary>
        /// ✔Fuente Negrita
        /// <br/><b>Sintaxis:</b>
        /// <br/><c>bold: [true | false]</c>
        /// </summary>
        public bool? bold { get; set; }
        //public string underline { get; set; }
        //public string underline_single { get; set; }
        //public string underline_double { get; set; }
        //public string underline_doubleaccount { get; set; }
        //public string strikethrough { get; set; }
        //public string text_decoration_style { get; set; }
        //public string text_decoration_line { get; set; }
        //public string text_transform { get; set; }
        //public string shrink_to_fit { get; set; }
        /// <summary>
        /// ✔La propiedad <c>text-align</c> establece la alineación horizontal del contenido a nivel de línea dentro de un elemento de bloque o caja de celda-tabla. Esto significa que funciona como vertical-align pero en dirección horizontal.<br/>
        /// <b>Sintaxis</b>
        /// <br/><c>text-align: [start | left | center | end | right | justify | fill | distributed | general];</c>
        /// </summary>
        public string text_align { get; set; }
        //public string text_rotation { get; set; }
        //public string text_orientation { get; set; }
        /// <summary>
        /// ✔La propiedad <c>text-wrap</c> controla cómo se ajusta el texto dentro de un elemento.
        /// <br/><b>Sintaxis</b>
        /// <br/><c>text-wrap: [wrap | true | nowrap | false];</c>
        /// </summary>
        public object text_wrap { get; set; }
        /// <summary>
        /// ✔La propiedad <c>vertical-align</c> especifica el alineado vertical de un elemento en línea o una celda de una tabla.
        /// <br/><b>Sintaxis</b>
        /// <br/><c>text-align: [top | middle | center | bottom | justify | fill | distributed];</c>
        /// </summary>
        public string vertical_align { get; set; }
        //public string writing_mode { get; set; }
        /// <summary>
        /// ✔Number-format es una propiedad que establece el formato de la celda.
        /// <br/>Se pueden utilizar los formatos de números integrados o several build (<b>int</b>, <a href="https://github.com/EPPlusSoftware/EPPlus/wiki/Formatting-and-styling">Ver Info</a>), o un formato personalizado (<b>string</b>, Ej. <c>$###,###,##0.00</c>)
        /// </summary>
        public string number_format { get; set; }
        public double? width { get; set; }
        public double? height { get; set; }

        #region OTHER CONFIGURATIONS
        /// <summary>
        /// ✔La propiedad <c>merge</c> combina el rango de celdas definidas.
        /// <br/>
        /// <br/><c><b>true</b></c> = Convinar la celda en el rango establecido
        /// <br/><c><b>false</b></c> = (default) No convinar la celda en el rango establecido (este repetirá el valor en cada celda del rango establecido)
        /// </summary>
        public bool? merge { get; set; }
        #endregion OTHER CONFIGURATIONS
    }
}
