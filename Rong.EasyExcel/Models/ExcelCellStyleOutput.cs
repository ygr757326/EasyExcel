using System;
using System.Reflection;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 单元格样式信息
    /// </summary>
    public class ExcelCellStyleOutput<TCellStyle, TStyle, TFont>
        where TStyle : Attribute
        where TFont : Attribute
    {
        /// <summary>
        /// 表头对应的字段属性
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 单元格样式
        /// </summary>
        public TCellStyle CellStyle { get; set; }

        /// <summary>
        /// 样式
        /// </summary>
        public TStyle StyleAttr { get; set; }

        /// <summary>
        /// 字体
        /// </summary>
        public TFont FontAttr { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelCellStyleOutput()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelCellStyleOutput(PropertyInfo propertyInfo, TCellStyle cellStyle, TStyle styleAttr, TFont fontAttr)
        {
            PropertyInfo = propertyInfo;
            CellStyle = cellStyle;
            StyleAttr = styleAttr;
            FontAttr = fontAttr;
        }
    }
}
