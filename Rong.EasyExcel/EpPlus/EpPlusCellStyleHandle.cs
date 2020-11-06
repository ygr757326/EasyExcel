using Rong.EasyExcel.Attributes;
using OfficeOpenXml.Style;

namespace Rong.EasyExcel.EpPlus
{
    /// <summary>
    /// EpPlus 单元格样式处理
    /// </summary>
    public class EpPlusCellStyleHandle : IEpPlusCellStyleHandle
    {
        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusCellStyleHandle()
        {
        }

        /// <summary>
        /// 设置表头单元格样式和字体
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="fontAttr"></param>
        /// <param name="styleAttr"></param>
        /// <returns></returns>
        public virtual void SetHeaderCellStyleAndFont(ExcelStyle cellStyle, HeaderStyleAttribute styleAttr, HeaderFontAttribute fontAttr)
        {
            //表头默认样式
            SetHeaderCellStyle(cellStyle, styleAttr);
            SetHeaderCellFont(cellStyle.Font, fontAttr);
        }

        /// <summary>
        /// 设置数据单元格样式和字体
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="styleAttr"></param>
        /// <param name="fontAttr"></param>
        /// <returns></returns>

        public virtual void SetDataCellStyleAndFont(ExcelStyle cellStyle, DataStyleAttribute styleAttr, DataFontAttribute fontAttr)
        {
            //表头默认样式
            SetDataCellStyle(cellStyle, styleAttr);
            SetDataCellFont(cellStyle.Font, fontAttr);
        }

        /// <summary>
        /// 设置表头单元格样式
        /// </summary>
        public virtual void SetHeaderCellStyle(ExcelStyle cellStyle, HeaderStyleAttribute styleAttr)
        {
            if (cellStyle == null)
            {
                return;
            }

            styleAttr = styleAttr ?? new HeaderStyleAttribute();

            if (!string.IsNullOrWhiteSpace(styleAttr.DataFormat)) cellStyle.Numberformat.Format = styleAttr.DataFormat;
            cellStyle.ShrinkToFit = styleAttr.ShrinkToFit;
            cellStyle.WrapText = styleAttr.WrapText;
            cellStyle.Hidden = styleAttr.IsHidden;
            cellStyle.Locked = styleAttr.IsLocked;
            cellStyle.WrapText = styleAttr.WrapText;

            if (styleAttr.Indention > -1) cellStyle.Indent = styleAttr.Indention;
            if (styleAttr.Rotation > -1) cellStyle.TextRotation = styleAttr.Rotation;
            if (styleAttr.Alignment > -1) cellStyle.HorizontalAlignment = (ExcelHorizontalAlignment)styleAttr.Alignment;
            if (styleAttr.VerticalAlignment >= -1) cellStyle.VerticalAlignment = (ExcelVerticalAlignment)styleAttr.VerticalAlignment;

            //边框样式、颜色
            if (styleAttr.BorderLeft > -1)
            {
                cellStyle.Border.Left.Style = (ExcelBorderStyle)styleAttr.BorderLeft;

                if (styleAttr.LeftBorderColor > -1) cellStyle.Border.Left.Color.Indexed = styleAttr.LeftBorderColor;
            }
            if (styleAttr.BorderRight > -1)
            {
                cellStyle.Border.Right.Style = (ExcelBorderStyle)styleAttr.BorderRight;

                if (styleAttr.RightBorderColor > -1) cellStyle.Border.Right.Color.Indexed = styleAttr.RightBorderColor;
            }
            if (styleAttr.BorderTop > -1)
            {
                cellStyle.Border.Top.Style = (ExcelBorderStyle)styleAttr.BorderTop;

                if (styleAttr.TopBorderColor > -1) cellStyle.Border.Top.Color.Indexed = styleAttr.TopBorderColor;
            }
            if (styleAttr.BorderBottom > -1)
            {
                cellStyle.Border.Bottom.Style = (ExcelBorderStyle)styleAttr.BorderBottom;

                if (styleAttr.BottomBorderColor > -1) cellStyle.Border.Bottom.Color.Indexed = styleAttr.BottomBorderColor;
            }

            //对角线
            //cellStyle.BorderDiagonal = (BorderDiagonal)styleAttr.BorderDiagonal;
            if (styleAttr.BorderDiagonalLineStyle > -1)
            {
                cellStyle.Border.Diagonal.Style = (ExcelBorderStyle)styleAttr.BorderDiagonalLineStyle;

                if (styleAttr.BorderDiagonalColor > -1) cellStyle.Border.Diagonal.Color.Indexed = styleAttr.BorderDiagonalColor;
            }

            //填充
            if (styleAttr.FillPattern > -1)
            {
                cellStyle.Fill.PatternType = (ExcelFillStyle)styleAttr.FillPattern;

                if (styleAttr.FillBackgroundColor > -1) cellStyle.Fill.BackgroundColor.Indexed = styleAttr.FillBackgroundColor;
            }
            //cellStyle.Fill.ForegroundColor = styleAttr.FillForegroundColor;
        }

        /// <summary>
        /// 设置表头单元格的字体
        /// </summary>
        public virtual void SetHeaderCellFont(ExcelFont font, HeaderFontAttribute fontAttr)
        {
            if (font == null)
            {
                return;
            }

            fontAttr = fontAttr ?? new HeaderFontAttribute();

            // font.FontHeight = fontAttr.FontHeight;
            // font.Charset = fontAttr.Charset;
            // font.TypeOffset = (FontSuperScript)fontAttr.TypeOffset;

            if (!string.IsNullOrWhiteSpace(fontAttr.FontName)) font.Name = fontAttr.FontName;
            font.Italic = fontAttr.IsItalic;
            font.Strike = fontAttr.IsStrikeout;
            font.Bold = fontAttr.IsBold;

            if (fontAttr.Underline > -1) font.UnderLineType = (ExcelUnderLineType)fontAttr.Underline;
            if (fontAttr.FontHeightInPoints > -1) font.Size = fontAttr.FontHeightInPoints;
            if (fontAttr.Color > -1) font.Color.Indexed = fontAttr.Color;
        }

        /// <summary>
        /// 设置数据单元格样式
        /// </summary>
        public virtual void SetDataCellStyle(ExcelStyle cellStyle, DataStyleAttribute styleAttr)
        {
            if (cellStyle == null)
            {
                return;
            }

            styleAttr = styleAttr ?? new DataStyleAttribute();

            if (!string.IsNullOrWhiteSpace(styleAttr.DataFormat)) cellStyle.Numberformat.Format = styleAttr.DataFormat;
            cellStyle.ShrinkToFit = styleAttr.ShrinkToFit;
            cellStyle.WrapText = styleAttr.WrapText;
            cellStyle.Hidden = styleAttr.IsHidden;
            cellStyle.Locked = styleAttr.IsLocked;
            cellStyle.WrapText = styleAttr.WrapText;

            if (styleAttr.Indention > -1) cellStyle.Indent = styleAttr.Indention;
            if (styleAttr.Rotation > -1) cellStyle.TextRotation = styleAttr.Rotation;
            if (styleAttr.Alignment > -1) cellStyle.HorizontalAlignment = (ExcelHorizontalAlignment)styleAttr.Alignment;
            if (styleAttr.VerticalAlignment >= -1) cellStyle.VerticalAlignment = (ExcelVerticalAlignment)styleAttr.VerticalAlignment;

            //边框样式、颜色
            if (styleAttr.BorderLeft > -1)
            {
                cellStyle.Border.Left.Style = (ExcelBorderStyle)styleAttr.BorderLeft;

                if (styleAttr.LeftBorderColor > -1) cellStyle.Border.Left.Color.Indexed = styleAttr.LeftBorderColor;
            }
            if (styleAttr.BorderRight > -1)
            {
                cellStyle.Border.Right.Style = (ExcelBorderStyle)styleAttr.BorderRight;

                if (styleAttr.RightBorderColor > -1) cellStyle.Border.Right.Color.Indexed = styleAttr.RightBorderColor;
            }
            if (styleAttr.BorderTop > -1)
            {
                cellStyle.Border.Top.Style = (ExcelBorderStyle)styleAttr.BorderTop;

                if (styleAttr.TopBorderColor > -1) cellStyle.Border.Top.Color.Indexed = styleAttr.TopBorderColor;
            }
            if (styleAttr.BorderBottom > -1)
            {
                cellStyle.Border.Bottom.Style = (ExcelBorderStyle)styleAttr.BorderBottom;

                if (styleAttr.BottomBorderColor > -1) cellStyle.Border.Bottom.Color.Indexed = styleAttr.BottomBorderColor;
            }

            //对角线
            //cellStyle.BorderDiagonal = (BorderDiagonal)styleAttr.BorderDiagonal;
            if (styleAttr.BorderDiagonalLineStyle > -1)
            {
                cellStyle.Border.Diagonal.Style = (ExcelBorderStyle)styleAttr.BorderDiagonalLineStyle;

                if (styleAttr.BorderDiagonalColor > -1) cellStyle.Border.Diagonal.Color.Indexed = styleAttr.BorderDiagonalColor;
            }

            //填充
            if (styleAttr.FillPattern > -1)
            {
                cellStyle.Fill.PatternType = (ExcelFillStyle)styleAttr.FillPattern;

                if (styleAttr.FillBackgroundColor > -1) cellStyle.Fill.BackgroundColor.Indexed = styleAttr.FillBackgroundColor;
            }
            //cellStyle.Fill.ForegroundColor = styleAttr.FillForegroundColor;

        }

        /// <summary>
        /// 设置数据单元格的字体
        /// </summary>
        public virtual void SetDataCellFont(ExcelFont font, DataFontAttribute fontAttr)
        {
            if (font == null)
            {
                return;
            }

            fontAttr = fontAttr ?? new DataFontAttribute();

            // font.FontHeight = fontAttr.FontHeight;
            // font.Charset = fontAttr.Charset;
            // font.TypeOffset = (FontSuperScript)fontAttr.TypeOffset;

            if (!string.IsNullOrWhiteSpace(fontAttr.FontName)) font.Name = fontAttr.FontName;
            font.Italic = fontAttr.IsItalic;
            font.Strike = fontAttr.IsStrikeout;
            font.Bold = fontAttr.IsBold;

            if (fontAttr.Underline > -1) font.UnderLineType = (ExcelUnderLineType)fontAttr.Underline;
            if (fontAttr.FontHeightInPoints > -1) font.Size = fontAttr.FontHeightInPoints;
            if (fontAttr.Color > -1) font.Color.Indexed = fontAttr.Color;
        }
    }
}
