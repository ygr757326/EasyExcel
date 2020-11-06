using Rong.EasyExcel.Attributes;
using OfficeOpenXml.Style;

namespace Rong.EasyExcel.EpPlus
{
    /// <summary>
    /// EpPlus 单元格样式处理
    /// </summary>
    public interface IEpPlusCellStyleHandle
    {
        /// <summary>
        /// 设置表头单元格样式和字体
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="fontAttr"></param>
        /// <param name="styleAttr"></param>
        /// <returns></returns>
        void SetHeaderCellStyleAndFont(ExcelStyle cellStyle, HeaderStyleAttribute styleAttr,
            HeaderFontAttribute fontAttr);

        /// <summary>
        /// 设置数据单元格样式和字体
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="styleAttr"></param>
        /// <param name="fontAttr"></param>
        /// <returns></returns>

        void SetDataCellStyleAndFont(ExcelStyle cellStyle, DataStyleAttribute styleAttr,
            DataFontAttribute fontAttr);

        /// <summary>
        /// 设置表头单元格样式
        /// </summary>
        void SetHeaderCellStyle(ExcelStyle cellStyle, HeaderStyleAttribute styleAttr);

        /// <summary>
        /// 设置表头单元格的字体
        /// </summary>
        void SetHeaderCellFont(ExcelFont font, HeaderFontAttribute fontAttr);

        /// <summary>
        /// 设置数据单元格样式
        /// </summary>
        void SetDataCellStyle(ExcelStyle cellStyle, DataStyleAttribute styleAttr);

        /// <summary>
        /// 设置数据单元格的字体
        /// </summary>
        void SetDataCellFont(ExcelFont font, DataFontAttribute fontAttr);
    }
}
