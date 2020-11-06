using Rong.EasyExcel.Models;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Rong.EasyExcel.EpPlus.Import
{
    /// <summary>
    /// EpPlus Excel 导入实现（版本号5.0.0之前的为免费版）
    /// </summary>
    public class EpPlusExcelImportBase : ExcelImportBase<ExcelWorkbook, ExcelWorksheet, ExcelRow, ExcelRange>
    {
        private readonly IEpPlusExcelHandle _epPlusExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusExcelImportBase(IEpPlusExcelHandle epPlusExcelHandle)
        {
            _epPlusExcelHandle = epPlusExcelHandle;
        }

        protected override ExcelWorkbook GetWorkbook(Stream fileStream)
        {
            var excelPackage = new ExcelPackage(fileStream);
            return excelPackage.Workbook;
        }

        protected override int GetWorksheetNumber(ExcelWorkbook workbook)
        {
            return workbook.Worksheets.Count;
        }

        protected override ExcelWorksheet GetWorksheet(ExcelWorkbook workbook, int sheetIndex)
        {
            return workbook.Worksheets[sheetIndex];
        }

        protected override string GetWorksheetName(ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            return worksheet.Name;
        }

        protected override ExcelRow GetHeaderRow(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelImportOptions options)
        {
            return worksheet.Row(options.HeaderRowIndex);
        }

        protected override List<ExcelHeaderCell> GetHeaderCells(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRow headerRow)
        {
            var headerCells = new List<ExcelHeaderCell>();

            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                var name = _epPlusExcelHandle.GetMergedCellValue(worksheet, headerRow.Row, i)?.ToString();

                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                headerCells.Add(new ExcelHeaderCell(name, headerRow.Row, i));
            }

            return headerCells;
        }

        protected override ExcelDataRowRangeIndex GetDataRowStartAndEndRowIndex(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelImportOptions options)
        {
            int startRowIndex = options.DataRowStartIndex - 1;
            int endRowIndex = worksheet.Dimension.End.Row - 1;
            if (options.DataRowEndIndex != null)
            {
                int end = (int)options.DataRowEndIndex - 1;
                endRowIndex = end > endRowIndex ? endRowIndex : end;
            }
            return new ExcelDataRowRangeIndex(startRowIndex, endRowIndex);
        }

        protected override ExcelRow GetDataRow(ExcelWorkbook workbook, ExcelWorksheet worksheet, int rowIndex)
        {
            return worksheet.Row(rowIndex + 1);
        }


        protected override object ConvertCellValue(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRow dataRow, int columnIndex, PropertyInfo property)
        {
            return _epPlusExcelHandle.ConverterCellValue(worksheet, dataRow.Row, columnIndex, property.PropertyType);
        }

        protected override string GetCellAddress(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRow dataRow, int columnIndex)
        {
            return _epPlusExcelHandle.GetCellAddress(dataRow.Row, columnIndex);
        }
    }
}
