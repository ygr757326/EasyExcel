using NPOI.SS.UserModel;
using System;
using System.IO;

namespace Rong.EasyExcel.Npoi
{
    /// <summary>
    /// Npoi Excel 处理接口
    /// </summary>
    public interface INpoiExcelHandle
    {
        /// <summary>
        /// 获取 IWorkbook
        /// </summary>
        /// <param name="physicalPath">物理路径</param>
        /// <returns></returns>
        IWorkbook GetWorkbook(string physicalPath);

        /// <summary>
        /// 获取 IWorkbook
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        IWorkbook GetWorkbook(Stream fileStream);

        /// <summary>
        /// 获取单元格合并信息
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <param name="cell">单元格</param>
        /// <returns>MergedInfo</returns>
        ExcelCellMergedInfo GetCellMergedInfo(ISheet sheet, ICell cell);

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        object GetCellValue(ICell cell);

        /// <summary>
        /// 得到公式单元格的值
        /// </summary>
        /// <param name="formulaValue"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        object GetCellValue(CellValue formulaValue, ICell cell);

        /// <summary>
        /// 获取所在合并的单元格区域的值
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <param name="cell">单元格</param>
        /// <returns>若是合并的单元格：返回合并区域的第一个值。若是非合并单元格：返回当前表格的值</returns>
        object GetMergedCellValue(ISheet sheet, ICell cell);

        /// <summary>
        /// 获取单元格值的默认格式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        short GetDefaultFormat(IWorkbook workbook,
            CellValueDefaultFormatEnum cellValue = CellValueDefaultFormatEnum.文本);

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">当前列下标,起始0</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        object ConverterCellValue(IRow row, int columnIndex, Type valueType);

        /// <summary>
        /// 获取默认的单元格样式
        /// <para>在.xls工作簿中最多可以定义4000种样式,所以该方法要在循环外定义</para>
        /// </summary>
        /// <param name="workbook">Workbook</param>
        /// <param name="style">CellStyleEnum</param>
        /// <returns></returns>
        ICellStyle GetDefaultCellStyle(IWorkbook workbook, ExcelCellStyleEnum style = ExcelCellStyleEnum.默认);

        /// <summary>
        /// 单列设置列宽（该方法必须在创建列后才能设置，不能在创建列前设置；列宽自动调整，必须有列数据才能处理））
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列下标（起始下表：0）</param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        /// <param name="columnAutoSize">是否自动调整</param>
        void SetColumnWidth(ISheet sheet, int columnIndex, int columnSize, bool columnAutoSize);

        /// <summary>
        /// 统一设置列宽（该方法必须在创建sheet后马上设置，不能在创建列后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        void SetColumnWidth(ISheet sheet, int columnSize);

        /// <summary>
        /// 单行设置行高（该方法必须在创建行后才能设置，不能在创建行前设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">行</param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        void SetRowHeight(ISheet sheet, IRow row, short rowHeight);

        /// <summary>
        /// 统一设置行高（该方法必须在创建sheet后马上设置，不能在创建行后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        void SetRowHeight(ISheet sheet, short rowHeight);

        /// <summary>
        /// IWorkbook写入流并转换为字节
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        byte[] GetAsByteArray(IWorkbook workbook);

        /// <summary>
        /// 合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow">起始行（起始下标：0）</param>
        /// <param name="lastRow">结束行（起始下标：0）</param>
        /// <param name="firstCol">起始列（起始下标：0）</param>
        /// <param name="lastCol">结束列（起始下标：0）</param>
        void MergedRegion(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol);

        /// <summary>
        /// 获取区域字符串（如： "A1"）
        /// </summary>
        /// <param name="rowIndex">行下标（起始下标：0）</param>
        /// <param name="columnIndex">列下标（起始下标：0）</param>
        string GetCellAddress(int rowIndex, int columnIndex);

        /// <summary>
        /// 获取区域字符串（如： "A1"）
        /// </summary>
        /// <param name="cell">单元格</param>
        string GetCellAddress(ICell cell);
    }
}
