using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using static EpplusTemplatesCS.ToolKits;

namespace EpplusTemplatesCS
{
    class Dictionaries
    {
        private static Dictionary<string, ExcelBorderStyle> border_style = new Dictionary<string, ExcelBorderStyle>() {
            {"default", ExcelBorderStyle.Medium },
            {"dash-dot", ExcelBorderStyle.DashDot },
            {"dash-dot-dot", ExcelBorderStyle.DashDotDot },

            {"dashed", ExcelBorderStyle.Dashed },   //dash
            {"dash", ExcelBorderStyle.Dashed },

            {"dotted", ExcelBorderStyle.Dotted },   //dot
            {"dot", ExcelBorderStyle.Dotted },
            {"double", ExcelBorderStyle.Double },
            {"hair", ExcelBorderStyle.Hair },

            {"medium", ExcelBorderStyle.Medium },   //single, solid
            {"single", ExcelBorderStyle.Medium },
            {"solid", ExcelBorderStyle.Medium },

            {"medium-dash-dot", ExcelBorderStyle.MediumDashDot },
            {"medium-dash-dot-dot", ExcelBorderStyle.MediumDashDotDot },
            {"medium-dashed", ExcelBorderStyle.MediumDashed },
            {"thick", ExcelBorderStyle.Thick },
            {"thin", ExcelBorderStyle.Thin },
        };
        private static Dictionary<string, Color> color = new Dictionary<string, Color>()
        {
            {"default", Color.White},       //default
            {"alice-blue", Color.AliceBlue},
            {"antique-white", Color.AntiqueWhite},
            {"aqua", Color.Aqua},
            {"aquamarine", Color.Aquamarine},
            {"azure", Color.Azure},
            {"beige", Color.Beige},
            {"bisque", Color.Bisque},
            {"black", Color.Black},
            {"blanched-almond", Color.BlanchedAlmond},
            {"blue", Color.Blue},
            {"blue-violet", Color.BlueViolet},
            {"brown", Color.Brown},
            {"burly-wood", Color.BurlyWood},
            {"cadet-blue", Color.CadetBlue},
            {"chartreuse", Color.Chartreuse},
            {"chocolate", Color.Chocolate},
            {"coral", Color.Coral},
            {"cornflower-blue", Color.CornflowerBlue},
            {"cornsilk", Color.Cornsilk},
            {"crimson", Color.Crimson},
            {"cyan", Color.Cyan},
            {"dark-blue", Color.DarkBlue},
            {"dark-cyan", Color.DarkCyan},
            {"dark-goldenrod", Color.DarkGoldenrod},
            {"dark-gray", Color.DarkGray},
            {"dark-green", Color.DarkGreen},
            {"dark-khaki", Color.DarkKhaki},
            {"dark-magenta", Color.DarkMagenta},
            {"dark-olive-green", Color.DarkOliveGreen},
            {"dark-orange", Color.DarkOrange},
            {"dark-orchid", Color.DarkOrchid},
            {"dark-red", Color.DarkRed},
            {"dark-salmon", Color.DarkSalmon},
            {"dark-sea-green", Color.DarkSeaGreen},
            {"dark-slate-blue", Color.DarkSlateBlue},
            {"dark-slate-gray", Color.DarkSlateGray},
            {"dark-turquoise", Color.DarkTurquoise},
            {"dark-violet", Color.DarkViolet},
            {"deep-pink", Color.DeepPink},
            {"deep-sky-blue", Color.DeepSkyBlue},
            {"dim-gray", Color.DimGray},
            {"dodger-blue", Color.DodgerBlue},
            {"firebrick", Color.Firebrick},
            {"floral-white", Color.FloralWhite},
            {"forest-green", Color.ForestGreen},
            {"fuchsia", Color.Fuchsia},
            {"gainsboro", Color.Gainsboro},
            {"ghost-white", Color.GhostWhite},
            {"gold", Color.Gold},
            {"goldenrod", Color.Goldenrod},
            {"gray", Color.Gray},
            {"green", Color.Green},
            {"green-yellow", Color.GreenYellow},
            {"honeydew", Color.Honeydew},
            {"hot-pink", Color.HotPink},
            {"indian-red", Color.IndianRed},
            {"indigo", Color.Indigo},
            {"ivory", Color.Ivory},
            {"khaki", Color.Khaki},
            {"lavender", Color.Lavender},
            {"lavender-blush", Color.LavenderBlush},
            {"lawn-green", Color.LawnGreen},
            {"lemon-chiffon", Color.LemonChiffon},
            {"light-blue", Color.LightBlue},
            {"light-coral", Color.LightCoral},
            {"light-cyan", Color.LightCyan},
            {"light-goldenrod-yellow", Color.LightGoldenrodYellow},
            {"light-gray", Color.LightGray},
            {"light-green", Color.LightGreen},
            {"light-pink", Color.LightPink},
            {"light-salmon", Color.LightSalmon},
            {"light-sea-green", Color.LightSeaGreen},
            {"light-sky-blue", Color.LightSkyBlue},
            {"light-slate-gray", Color.LightSlateGray},
            {"light-steel-blue", Color.LightSteelBlue},
            {"light-yellow", Color.LightYellow},
            {"lime", Color.Lime},
            {"lime-green", Color.LimeGreen},
            {"linen", Color.Linen},
            {"magenta", Color.Magenta},
            {"maroon", Color.Maroon},
            {"medium-aquamarine", Color.MediumAquamarine},
            {"medium-blue", Color.MediumBlue},
            {"medium-orchid", Color.MediumOrchid},
            {"medium-purple", Color.MediumPurple},
            {"medium-sea-green", Color.MediumSeaGreen},
            {"medium-slate-blue", Color.MediumSlateBlue},
            {"medium-spring-green", Color.MediumSpringGreen},
            {"medium-turquoise", Color.MediumTurquoise},
            {"medium-violet-red", Color.MediumVioletRed},
            {"midnight-blue", Color.MidnightBlue},
            {"mint-cream", Color.MintCream},
            {"misty-rose", Color.MistyRose},
            {"moccasin", Color.Moccasin},
            {"navajo-white", Color.NavajoWhite},
            {"navy", Color.Navy},
            {"old-lace", Color.OldLace},
            {"olive", Color.Olive},
            {"olive-drab", Color.OliveDrab},
            {"orange", Color.Orange},
            {"orange-red", Color.OrangeRed},
            {"orchid", Color.Orchid},
            {"pale-goldenrod", Color.PaleGoldenrod},
            {"pale-green", Color.PaleGreen},
            {"pale-turquoise", Color.PaleTurquoise},
            {"pale-violet-red", Color.PaleVioletRed},
            {"papaya-whip", Color.PapayaWhip},
            {"peach-puff", Color.PeachPuff},
            {"peru", Color.Peru},
            {"pink", Color.Pink},
            {"plum", Color.Plum},
            {"powder-blue", Color.PowderBlue},
            {"purple", Color.Purple},
            //{"rebecca-purple", Color.RebeccaPurple},
            {"red", Color.Red},
            {"rosy-brown", Color.RosyBrown},
            {"royal-blue", Color.RoyalBlue},
            {"saddle-brown", Color.SaddleBrown},
            {"salmon", Color.Salmon},
            {"sandy-brown", Color.SandyBrown},
            {"sea-green", Color.SeaGreen},
            {"sea-shell", Color.SeaShell},
            {"sienna", Color.Sienna},
            {"silver", Color.Silver},
            {"sky-blue", Color.SkyBlue},
            {"slate-blue", Color.SlateBlue},
            {"slate-gray", Color.SlateGray},
            {"snow", Color.Snow},
            {"spring-green", Color.SpringGreen},
            {"steel-blue", Color.SteelBlue},
            {"tan", Color.Tan},
            {"teal", Color.Teal},
            {"thistle", Color.Thistle},
            {"tomato", Color.Tomato},
            {"turquoise", Color.Turquoise},
            {"violet", Color.Violet},
            {"wheat", Color.Wheat},
            {"white", Color.White},
            {"white-smoke", Color.WhiteSmoke},
            {"yellow", Color.Yellow},
            {"yellow-green", Color.YellowGreen},
        };

        public static Dictionary<string, ExcelFillStyle> pattern_type = new Dictionary<string, ExcelFillStyle>()
        {
            {"default", ExcelFillStyle.Solid },
            {"dark-down", ExcelFillStyle.DarkDown },
            {"dark-gray", ExcelFillStyle.DarkGray },
            {"dark-grid", ExcelFillStyle.DarkGrid },
            {"dark-horizontal", ExcelFillStyle.DarkHorizontal },
            {"dark-trellis", ExcelFillStyle.DarkTrellis },
            {"dark-up", ExcelFillStyle.DarkUp },
            {"dark-vertical", ExcelFillStyle.DarkVertical },
            {"gray_0625", ExcelFillStyle.Gray0625 },
            {"gray-125", ExcelFillStyle.Gray125 },
            {"light-down", ExcelFillStyle.LightDown },
            {"light-gray", ExcelFillStyle.LightGray },
            {"light-grid", ExcelFillStyle.LightGrid },
            {"light-horizontal", ExcelFillStyle.LightHorizontal },
            {"light-trellis", ExcelFillStyle.LightTrellis },
            {"light-up", ExcelFillStyle.LightUp },
            {"light-vertical", ExcelFillStyle.LightVertical },
            {"medium-gray", ExcelFillStyle.MediumGray },
            {"solid", ExcelFillStyle.Solid },
        };
        private static Dictionary<string, ExcelHorizontalAlignment> text_align = new Dictionary<string, ExcelHorizontalAlignment>() {
            { "default", ExcelHorizontalAlignment.General},     //default

            { "left", ExcelHorizontalAlignment.Left}, //start
            { "start", ExcelHorizontalAlignment.Left},

            { "center", ExcelHorizontalAlignment.Center},
            { "center-continuous", ExcelHorizontalAlignment.CenterContinuous},

            { "right", ExcelHorizontalAlignment.Right},   //end
            { "end", ExcelHorizontalAlignment.Right},

            { "justify", ExcelHorizontalAlignment.Justify},
            { "fill", ExcelHorizontalAlignment.Fill},
            { "distributed", ExcelHorizontalAlignment.Distributed},
            { "general", ExcelHorizontalAlignment.General}
        };
        public static Dictionary<string, ExcelVerticalAlignment> vertical_align = new Dictionary<string, ExcelVerticalAlignment>()
        {
            {"top", ExcelVerticalAlignment.Top },

            {"center", ExcelVerticalAlignment.Center },     //middle
            {"middle", ExcelVerticalAlignment.Center },

            {"bottom", ExcelVerticalAlignment.Bottom },
            {"distributed", ExcelVerticalAlignment.Distributed },
            {"justify", ExcelVerticalAlignment.Justify }
        };

        #region MÉTODOS
        /// <summary>
        /// Obtiene un objecto Color a partir de su nombre
        /// <br/>Nomenclatura aceptada (Ej.):
        /// <br/><c>yellow | yellow-green | Yellow | YellowGreen</c>
        /// </summary>
        /// <param name="color_name">Nombre del color</param>
        /// <returns>Color</returns>
        public static Color GetColor(string color_name)
        {
            color_name = Regex.IsMatch(color_name, "[A-Z]") ? Text.RenameColor(color_name) : color_name.ToLower();  //Formas de nombres permitidos: yellow | yellow-green | Yellow | YellowGreen

            Color _color = color["default"]; //por defecto

            if (color.ContainsKey(color_name))
            {
                _color = color[color_name];
            }

            return _color;
        }

        /// <summary>
        /// Obtiene un objeto ExcelBorderStyle (Estilo de Borde) a partir de su nombre
        /// <br/>Nomenclatura aceptada (Ej.):
        /// <br/><c>center | center-continuous | Center | CenterContinuous</c>
        /// </summary>
        /// <param name="border_style_name">Nombre del estilo de borde</param>
        /// <returns>ExcelBorderStyle</returns>
        public static ExcelBorderStyle GetBorderStyle(string border_style_name)
        {
            border_style_name = Regex.IsMatch(border_style_name, "[A-Z]") ? Text.RenameColor(border_style_name) : border_style_name.ToLower();  //Formas de nombres permitidos: dash | dash-dot | Dash | DashDot

            ExcelBorderStyle _border_style = border_style["default"]; //por defecto

            if (border_style.ContainsKey(border_style_name))
            {
                _border_style = border_style[border_style_name];
            }

            return _border_style;
        }

        /// <summary>
        /// Obtiene un objecto ExcelHorizontalAlignment (Alineación de texto de manera horizontal) a partir de su nombre
        /// </summary>
        /// <param name="text_align_name"></param>
        /// <returns></returns>
        public static ExcelHorizontalAlignment GetTextAlign(string text_align_name)
        {
            text_align_name = Regex.IsMatch(text_align_name, "[A-Z]") ? Text.RenameColor(text_align_name) : text_align_name.ToLower();  //Formas de nombres permitidos: dash | dash-dot | Dash | DashDot

            ExcelHorizontalAlignment _text_align = text_align["default"]; //por defecto

            if (text_align.ContainsKey(text_align_name))
            {
                _text_align = text_align[text_align_name];
            }

            return _text_align;
        }

        #endregion MÉTODOS

        #region OTROS MÉTODOS

        #endregion OTROS MÉTODOS
    }
}
