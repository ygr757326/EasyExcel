using Rong.EasyExcel.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Rong.EasyExcel.Npoi.Import
{
    /// <summary>
    /// Npoi Excel 导入实现
    /// </summary>
    public class NpoiExcelImportBase : ExcelImportBase<IWorkbook, ISheet, IRow, ICell>
    {
        public INpoiExcelHandle _npoiExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public NpoiExcelImportBase(INpoiExcelHandle npoiExcelHandle)
        {
            _npoiExcelHandle = npoiExcelHandle;
        }

        protected override IWorkbook GetWorkbook(Stream fileStream)
        {
            return _npoiExcelHandle.GetWorkbook(fileStream);
        }
        protected override int GetWorksheetNumber(IWorkbook workbook)
        {
            return workbook.NumberOfSheets;
        }

        protected override ISheet GetWorksheet(IWorkbook workbook, int sheetIndex)
        {
            return workbook.GetSheetAt(sheetIndex);
        }

        protected override string GetWorksheetName(IWorkbook workbook, ISheet worksheet)
        {
            return worksheet.SheetName;
        }

        protected override IRow GetHeaderRow(IWorkbook workbook, ISheet worksheet, ExcelImportOptions options)
        {
            return worksheet.GetRow(options.HeaderRowIndex - 1);
        }

        protected override List<ExcelHeaderCell> GetHeaderCells(IWorkbook workbook, ISheet worksheet, IRow headerRow)
        {
            if (headerRow == null || headerRow.Cells.Count == 0)
            {
                throw new Exception($"工作表【{worksheet.SheetName}】表头行不能为空");
            }

            var headerCells = new List<ExcelHeaderCell>();

            foreach (var cell in headerRow.Cells)
            {
                string name = _npoiExcelHandle.GetMergedCellValue(worksheet, cell)?.ToString();
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }
                headerCells.Add(new ExcelHeaderCell(name, cell.RowIndex, cell.ColumnIndex));
            }

            return headerCells;
        }

        protected override ExcelDataRowRangeIndex GetDataRowStartAndEndRowIndex(IWorkbook workbook, ISheet worksheet, ExcelImportOptions options)
        {
            int startRowIndex = options.DataRowStartIndex - 1;
            int endRowIndex = worksheet.LastRowNum;
            if (options.DataRowEndIndex != null)
            {
                int end = (int)options.DataRowEndIndex - 1;
                endRowIndex = end > endRowIndex ? endRowIndex : end;
            }
            return new ExcelDataRowRangeIndex(startRowIndex, endRowIndex);
        }

        protected override IRow GetDataRow(IWorkbook workbook, ISheet worksheet, int rowIndex)
        {
            return worksheet.GetRow(rowIndex);
        }

        protected override object ConvertCellValue(IWorkbook workbook, ISheet worksheet, IRow dataRow, int columnIndex, PropertyInfo property)
        {
            return _npoiExcelHandle.ConverterCellValue(dataRow, columnIndex, property.PropertyType);
        }

        protected override string GetCellAddress(IWorkbook workbook, ISheet worksheet, IRow dataRow, int columnIndex)
        {
            return _npoiExcelHandle.GetCellAddress(dataRow.RowNum, columnIndex);
        }
    }
}
