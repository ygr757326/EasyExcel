using Rong.EasyExcel.Attributes;
using NPOI.SS.UserModel;

namespace Rong.EasyExcel.Npoi
{
    /// <summary>
    /// Npoi 单元格样式处理
    /// </summary>
    public interface INpoiCellStyleHandle
    {
        /// <summary>
        /// 设置表头单元格样式和字体
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontAttr"></param>
        /// <param name="styleAttr"></param>
        /// <returns></returns>
        ICellStyle SetHeaderCellStyleAndFont(IWorkbook workbook, HeaderStyleAttribute styleAttr,
            HeaderFontAttribute fontAttr);


        /// <summary>
        /// 设置数据单元格样式和字体
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="styleAttr"></param>
        /// <param name="fontAttr"></param>
        /// <returns></returns>

        ICellStyle SetDataCellStyleAndFont(IWorkbook workbook, DataStyleAttribute styleAttr,
            DataFontAttribute fontAttr);

        /// <summary>
        /// 创建表头单元格样式
        /// </summary>
        ICellStyle CreateHeaderCellStyle(IWorkbook workbook, HeaderStyleAttribute styleAttr);

        /// <summary>
        /// 创建表头单元格的字体
        /// </summary>
        IFont CreateHeaderCellFont(IWorkbook workbook, HeaderFontAttribute fontAttr);

        /// <summary>
        /// 创建数据单元格样式
        /// </summary>
        ICellStyle CreateDataCellStyle(IWorkbook workbook, DataStyleAttribute styleAttr);

        /// <summary>
        /// 创建数据单元格的字体
        /// </summary>
        IFont CreateDataCellFont(IWorkbook workbook, DataFontAttribute fontAttr);
    }
}
