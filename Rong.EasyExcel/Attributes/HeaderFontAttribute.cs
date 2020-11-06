using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using OfficeOpenXml.Style;
namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel表头特性（导出时用）
    /// <para>1.应用在类、字段、属性上，仅对表头有效</para>
    /// <para>2.若类和属性上都存在，则属性上的有效，类上的无效</para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property, AllowMultiple = false)]
    public sealed class HeaderFontAttribute : Attribute
    {
        /// <summary>
        /// 颜色索引编号
        /// <para>NPOI：<see cref="IndexedColors"/> <see cref="HSSFColor"/>，如：HSSFColor.Black.Index,IndexedColors.Black.Index</para>
        /// <para>EpPlus：<see cref="System.Drawing.Color"/></para>
        /// </summary>
        public short Color { get; set; } = -1;

        /// <summary>
        /// 字号大小
        /// <para>NPOI</para>
        /// <para>EpPlus</para>
        /// </summary>
        public short FontHeightInPoints { get; set; } = -1;

        /// <summary>
        /// 字体名称
        /// <para>NPOI</para>
        /// <para>EpPlus</para>
        /// </summary>
        public string FontName { get; set; }
        /// <summary>
        /// 字体高
        /// <para>NPOI</para>
        /// </summary>
        public double FontHeight { get; set; } = -1;
        /// <summary>
        /// 是否斜体
        /// <para>NPOI</para>
        /// <para>EpPlus</para>
        /// </summary>
        public bool IsItalic { get; set; }
        /// <summary>
        /// 是否有删除线
        /// <para>NPOI</para>
        /// <para>EpPlus</para>
        /// </summary>
        public bool IsStrikeout { get; set; }
        /// <summary>
        /// 字体上标下标
        /// <para>NPOI：<see cref="FontSuperScript"/></para>
        /// </summary>
        public short TypeOffset { get; set; } = -1;
        /// <summary>
        /// 下划线类型
        /// <para>NPOI：<see cref="FontUnderlineType"/></para>
        /// <para>EpPlus：<see cref="ExcelUnderLineType"/></para>
        /// </summary>
        public short Underline { get; set; } = -1;
        /// <summary>
        /// 字符集
        /// <para>NPOI</para>
        /// </summary>
        public short Charset { get; set; } = -1;
        /// <summary>
        /// 是否粗体
        /// <para>NPOI</para>
        /// <para>EpPlus</para>
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public HeaderFontAttribute()
        {
            IsBold = true;
        }
    }
}
