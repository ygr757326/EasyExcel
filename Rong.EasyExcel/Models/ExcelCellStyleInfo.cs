using System;
using System.Reflection;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 单元格样式信息
    /// </summary>
    public class ExcelCellStyleInfo<TStyle, TFont>
        where TStyle : Attribute
        where TFont : Attribute
    {
        /// <summary>
        /// 表头对应的字段属性
        /// </summary>
        public MemberInfo PropertyInfo { get; set; }

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
        public ExcelCellStyleInfo()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelCellStyleInfo(MemberInfo propertyInfo, TStyle styleAttr, TFont fontAttr)
        {
            PropertyInfo = propertyInfo;
            StyleAttr = styleAttr;
            FontAttr = fontAttr;
        }
    }
}
