using Rong.EasyExcel.Attributes;
using Rong.EasyExcel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Globalization;

namespace Rong.EasyExcel.EpPlus.Export
{
    /// <summary>
    /// EpPlus Excel 导入实现（版本号5.0.0之前的为免费版）
    /// </summary>
    public class EpPlusExcelExportBase : ExcelExportBase<ExcelWorkbook, ExcelWorksheet, ExcelRow, ExcelRange, ExcelStyle>
    {
        private readonly IEpPlusCellStyleHandle _epPlusCellStyleHandle;
        private readonly IEpPlusExcelHandle _epPlusExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusExcelExportBase(IEpPlusCellStyleHandle epPlusCellStyleHandle, IEpPlusExcelHandle epPlusExcelHandle)
        {
            _epPlusCellStyleHandle = epPlusCellStyleHandle;
            _epPlusExcelHandle = epPlusExcelHandle;
        }

        protected override ExcelWorkbook GetWorkbook(ExcelExportOptions options)
        {
            var excelPackage = new ExcelPackage();
            return excelPackage.Workbook;
        }

        protected override ExcelWorksheet CreateSheet(ExcelWorkbook workbook, ExcelExportOptions options)
        {
            return workbook.Worksheets.Add(options.SheetName);
        }

        protected override ExcelRange CreateCell(ExcelWorkbook workbook, ExcelWorksheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.Cells[rowIndex + 1, columnIndex + 1];
        }

        protected override void SetCellValue(ExcelWorkbook workbook, ExcelWorksheet sheet, ExcelRange cell, Type valueType, object value)
        {
            if (valueType.IsDateTime())
            {
                var date = value.GetTypedCellValue<DateTime>();
                if (date == default)
                {
                    cell.Value = date.ToString(CultureInfo.CurrentCulture);
                }
                else
                {
                    cell.Value = date;
                }
            }
            else if (valueType.IsTimeSpan())
            {
                cell.Value = value.GetTypedCellValue<DateTime>().ToString(CultureInfo.CurrentCulture);
            }
            else
            {
                cell.Value = value;
            }
        }

        protected override ExcelStyle CreateHeaderStyleAndFont<TExportDto>(ExcelWorkbook workbook, ExcelWorksheet worksheet,
            HeaderStyleAttribute styleAttr, HeaderFontAttribute fontAttr)
        {
            return null;
        }

        protected override ExcelStyle CreateDataStyleAndFont<TExportDto>(ExcelWorkbook workbook, ExcelWorksheet worksheet,
            DataStyleAttribute styleAttr, DataFontAttribute fontAttr)
        {
            return null;
        }

        protected override void SetHeaderCellStyleAndFont<TExportDto>(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRange cell,
            ExcelCellStyleOutput<ExcelStyle, HeaderStyleAttribute, HeaderFontAttribute> cellStyleInfo)
        {
            _epPlusCellStyleHandle.SetHeaderCellStyleAndFont(cell.Style, cellStyleInfo.StyleAttr, cellStyleInfo.FontAttr);

        }

        protected override void SetDataCellStyleAndFont<TExportDto>(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRange cell,
            ExcelCellStyleOutput<ExcelStyle, DataStyleAttribute, DataFontAttribute> cellStyleInfo)
        {
            _epPlusCellStyleHandle.SetDataCellStyleAndFont(cell.Style, cellStyleInfo.StyleAttr, cellStyleInfo.FontAttr);
        }


        protected override void SetColumnWidth(ExcelWorkbook workbook, ExcelWorksheet sheet, int columnIndex, int columnSize, bool columnAutoSize)
        {
            _epPlusExcelHandle.SetColumnWidth(sheet, columnIndex + 1, columnSize, columnAutoSize);
        }

        protected override void SetRowHeight(ExcelWorkbook workbook, ExcelWorksheet worksheet, int rowIndex, short rowHeight)
        {
            _epPlusExcelHandle.SetRowHeight(worksheet, rowIndex + 1, rowHeight);
        }

        protected override void SetMergedRegion(ExcelWorkbook workbook, ExcelWorksheet worksheet, int fromRowIndex, int toRowIndex,
            int fromColumnIndex, int toColumnIndex)
        {
            _epPlusExcelHandle.MergedRegion(worksheet, fromRowIndex + 1, toRowIndex + 1, fromColumnIndex + 1, toColumnIndex + 1);
        }

        protected override string GetCellAddress(ExcelWorkbook workbook, ExcelWorksheet worksheet, int rowIndex, int columnIndex)
        {
            return _epPlusExcelHandle.GetCellAddress(rowIndex + 1, columnIndex + 1);
        }


        protected override void SetCellFormula(ExcelWorkbook workbook, ExcelWorksheet worksheet, ExcelRange cell, string cellFormula)
        {
            cell.Formula = cellFormula;
        }

        protected override byte[] GetAsByteArray(ExcelWorkbook workbook, ExcelWorksheet sheet)
        {
            return _epPlusExcelHandle.GetAsByteArray(workbook, sheet);
        }
    }
}
