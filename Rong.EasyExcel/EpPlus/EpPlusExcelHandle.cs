using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;

namespace Rong.EasyExcel.EpPlus
{
    /// <summary>
    /// EpPlus excel 处理
    /// </summary>
    public class EpPlusExcelHandle : IEpPlusExcelHandle
    {
        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusExcelHandle()
        {
        }

        /// <summary>
        /// 获取 ExcelWorkbook
        /// </summary>
        /// <param name="physicalPath">物理路径</param>
        /// <returns></returns>
        public virtual ExcelWorkbook GetWorkbook(string physicalPath)
        {
            ExcelHelper.ValidationExcel(physicalPath);

            using (var stream = new FileStream(physicalPath, FileMode.Open, FileAccess.Read))
            {
                return GetWorkbook(stream);
            }
        }

        /// <summary>
        /// 获取 ExcelWorkbook
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        public virtual ExcelWorkbook GetWorkbook(Stream fileStream)
        {
            try
            {
                var excelPackage = new ExcelPackage(fileStream);

                return excelPackage.Workbook;
            }
            catch (Exception e)
            {
                throw new Exception($"获取 Workbook 出错：{e.Message}", e);
            }
        }

        /// <summary>
        /// 获取合并单元格的值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">当前行编号（起始下标：1）</param>
        /// <param name="columnIndex">当前列编号（起始下标：1）</param>
        /// <returns></returns>
        public virtual object GetMergedCellValue(ExcelWorksheet sheet, int rowIndex, int columnIndex)
        {
            ExcelRange cell = sheet.Cells[rowIndex, columnIndex];
            return GetMergedCellValue(sheet, cell);
        }

        /// <summary>
        /// 获取合并单元格的值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public virtual object GetMergedCellValue(ExcelWorksheet sheet, ExcelRange cell)
        {
            string address = sheet.MergedCells[cell.Start.Row, cell.Start.Column];
            if (address != null)
            {
                var excelAddress = new ExcelAddress(address);
                cell = sheet.Cells[excelAddress.Start.Row, excelAddress.Start.Column];
            }
            return cell.Value;
        }

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">当前行编号（起始下标：1）</param>
        /// <param name="columnIndex">当前列编号（起始下标：1）</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        public virtual object ConverterCellValue(ExcelWorksheet sheet, int rowIndex, int columnIndex, Type valueType)
        {
            ExcelRange cell = sheet.Cells[rowIndex, columnIndex];
            return ConverterCellValue(sheet, cell, valueType);
        }

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        public virtual object ConverterCellValue(ExcelWorksheet sheet, ExcelRange cell, Type valueType)
        {
            object cellValue = GetMergedCellValue(sheet, cell);
            if (string.IsNullOrWhiteSpace(cellValue?.ToString()))
            {
                return cellValue;
            }
            if (valueType.IsDateTime())
            {
                cellValue = cellValue.GetTypedCellValue<DateTime>();
            }
            else if (valueType.IsTimeSpan())
            {
                cellValue = cellValue.GetTypedCellValue<TimeSpan>();
            }
            return TypeDescriptor.GetConverter(valueType).ConvertFromInvariantString(cellValue.ToString());
        }

        /// <summary>
        /// 单列设置列宽（该方法必须在创建列后才能设置，不能在创建列前设置；列宽自动调整，必须有列数据才能处理））
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列下标（起始下标：1）</param>
        /// <param name="columnSize">列尺寸</param>
        /// <param name="columnAutoSize">是否自动调整</param>
        public virtual void SetColumnWidth(ExcelWorksheet sheet, int columnIndex, int columnSize, bool columnAutoSize)
        {
            if (columnSize > 0)
            {
                //创建“列”后才能设置，不能在创建“列”前设置
                sheet.Column(columnIndex).Width = columnSize > 255 ? 255 : columnSize;
            }
            else if (columnAutoSize)
            {
                //列宽自动调整，必须有“列数据”才能处理）
                sheet.Column(columnIndex).AutoFit();
            }
        }

        /// <summary>
        /// 统一设置列宽（该方法必须在创建sheet后马上设置，不能在创建列后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        public virtual void SetColumnWidth(ExcelWorksheet sheet, int columnSize)
        {
            if (columnSize > 0)
            {
                //必须在创建“sheet”后、创建“列“前设置，不能在创建列后才设置
                sheet.DefaultColWidth = columnSize > 255 ? 255 : columnSize;
            }
        }

        /// <summary>
        /// 单行设置行高（该方法必须在创建行后才能设置，不能在创建行前设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行下标（起始下标：1）</param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        public virtual void SetRowHeight(ExcelWorksheet sheet, int rowIndex, short rowHeight)
        {
            if (rowHeight > 0)
            {
                //必须在创建“行”后才能设置，不能在创建“行”前设置
                sheet.Row(rowIndex).Height = rowHeight > 409 ? 409 : rowHeight;
            }
        }

        /// <summary>
        /// 统一设置行高（该方法必须在创建sheet后马上设置，不能在创建行后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        public virtual void SetRowHeight(ExcelWorksheet sheet, short rowHeight)
        {
            if (rowHeight > 0)
            {
                //必须在创建“sheet”后马上设置，不能在创建“行”后才设置
                sheet.DefaultRowHeight = rowHeight > 409 ? 409 : rowHeight;
            }
        }

        /// <summary>
        /// 创建工作册并转换字节
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public virtual byte[] GetAsByteArray(ExcelWorkbook workbook, ExcelWorksheet sheet)
        {
            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Worksheets.Add(sheet.Name, sheet);

                return excelPackage.GetAsByteArray();
            }
        }

        /// <summary>
        /// 合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRow">起始行（起始下标：1）</param>
        /// <param name="toRow">结束行（起始下标：1）</param>
        /// <param name="fromColumn">起始列（起始下标：1）</param>
        /// <param name="toColumn">结束列（起始下标：1）</param>
        public virtual void MergedRegion(ExcelWorksheet sheet, int fromRow, int toRow, int fromColumn, int toColumn)
        {
            sheet.Cells[fromRow, fromColumn, toRow, toColumn].Merge = true;
        }

        /// <summary>
        /// 获取区域字符串（如： "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" ）
        /// </summary>
        /// <param name="fromRow">起始行（起始下标：1）</param>
        /// <param name="toRow">结束行（起始下标：1）</param>
        /// <param name="fromColumn">起始列（起始下标：1）</param>
        /// <param name="toColumn">结束列（起始下标：1）</param>
        public virtual string GetCellAddress(int fromRow, int toRow, int fromColumn, int toColumn)
        {
            return new ExcelAddress(fromRow, fromColumn, toRow, toColumn).Address;
        }

        /// <summary>
        /// 获取区域字符串（如： "A1" ）
        /// </summary>
        /// <param name="row">行下标（起始下标：1）</param>
        /// <param name="column">列下标（起始下标：1）</param>
        public virtual string GetCellAddress(int row,int column)
        {
            return new ExcelAddress(row, column, row, column).Address;
        }
    }
}
