using OfficeOpenXml;
using System;
using System.IO;

namespace Rong.EasyExcel.EpPlus
{
    /// <summary>
    /// EpPlus excel 处理
    /// </summary>
    public interface IEpPlusExcelHandle
    {
        /// <summary>
        /// 获取 ExcelWorkbook
        /// </summary>
        /// <param name="physicalPath">物理路径</param>
        /// <returns></returns>
        ExcelWorkbook GetWorkbook(string physicalPath);

        /// <summary>
        /// 获取 ExcelWorkbook
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        ExcelWorkbook GetWorkbook(Stream fileStream);

        /// <summary>
        /// 获取合并单元格的值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="row">当前行编号（起始下标：1）</param>
        /// <param name="column">当前列编号（起始下标：1）</param>
        /// <returns></returns>
        object GetMergedCellValue(ExcelWorksheet sheet, int row, int column);

        /// <summary>
        /// 获取合并单元格的值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        object GetMergedCellValue(ExcelWorksheet sheet, ExcelRange cell);

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">当前行（起始下标：1）</param>
        /// <param name="column">当前列（起始下标：1）</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        object ConverterCellValue(ExcelWorksheet sheet, int row, int column, Type valueType);

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        object ConverterCellValue(ExcelWorksheet sheet, ExcelRange cell, Type valueType);

        /// <summary>
        /// 单列设置列宽（该方法必须在创建列后才能设置，不能在创建列前设置；列宽自动调整，必须有列数据才能处理））
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列下标（起始下标：1）</param>
        /// <param name="columnSize">列尺寸</param>
        /// <param name="columnAutoSize">是否自动调整</param>
        void SetColumnWidth(ExcelWorksheet sheet, int columnIndex, int columnSize, bool columnAutoSize);

        /// <summary>
        /// 统一设置列宽（该方法必须在创建sheet后马上设置，不能在创建列后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        void SetColumnWidth(ExcelWorksheet sheet, int columnSize);

        /// <summary>
        /// 单行设置行高（该方法必须在创建行后才能设置，不能在创建行前设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行下标（起始下标：1）</param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        void SetRowHeight(ExcelWorksheet sheet, int rowIndex, short rowHeight);

        /// <summary>
        /// 统一设置行高（该方法必须在创建sheet后马上设置，不能在创建行后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        void SetRowHeight(ExcelWorksheet sheet, short rowHeight);

        /// <summary>
        /// 创建工作册并转换字节
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        byte[] GetAsByteArray(ExcelWorkbook workbook, ExcelWorksheet sheet);

        /// <summary>
        /// 合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRow">起始行（起始下标：1）</param>
        /// <param name="toRow">结束行（起始下标：1）</param>
        /// <param name="fromColumn">起始列（起始下标：1）</param>
        /// <param name="toColumn">结束列（起始下标：1）</param>
        void MergedRegion(ExcelWorksheet sheet, int fromRow, int toRow, int fromColumn, int toColumn);

        /// <summary>
        /// 获取区域字符串（如： "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" ）
        /// </summary>
        /// <param name="fromRow">起始行（起始下标：1）</param>
        /// <param name="toRow">结束行（起始下标：1）</param>
        /// <param name="fromColumn">起始列（起始下标：1）</param>
        /// <param name="toColumn">结束列（起始下标：1）</param>
        string GetCellAddress(int fromRow, int toRow, int fromColumn, int toColumn);

        /// <summary>
        /// 获取区域字符串（如： "A1" ）
        /// </summary>
        /// <param name="row">行下标（起始下标：1）</param>
        /// <param name="column">列下标（起始下标：1）</param>
        string GetCellAddress(int row, int column);
    }
}
