using Rong.EasyExcel.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.SS.Formula.Constant;

namespace Rong.EasyExcel.Npoi
{
    /// <summary>
    /// NPOI excel 处理
    /// </summary>
    public class NpoiExcelHandle : INpoiExcelHandle
    {
        /// <summary>
        /// 构造
        /// </summary>
        public NpoiExcelHandle()
        {
        }

        /// <summary>
        /// 获取 IWorkbook
        /// </summary>
        /// <param name="physicalPath">物理路径</param>
        /// <returns></returns>
        public virtual IWorkbook GetWorkbook(string physicalPath)
        {
            ExcelHelper.ValidationExcel(physicalPath);

            using (var stream = new FileStream(physicalPath, FileMode.Open, FileAccess.Read))
            {
                return GetWorkbook(stream);
            }
        }

        /// <summary>
        /// 获取 IWorkbook
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        public virtual IWorkbook GetWorkbook(Stream fileStream)
        {
            try
            {
                IWorkbook workbook = WorkbookFactory.Create(fileStream); ;

                //try
                //{
                //    // XSSFWorkbook：Excel >= 2007  的版本： .xlsx【1048576行,16384列】
                //    // SXSSFWorkbook：Excel >= 2007  的版本： .xlsx【硬盘空间换内存，适合处理大数据】
                //    workbook = new XSSFWorkbook(fileStream);
                //}
                //catch
                //{
                //    fileStream.Position = 0;
                //    // Excel <= 2003版 的版本：  .xls【65535行，256列】
                //    workbook = new HSSFWorkbook(fileStream);

                //}

                return workbook;
            }
            catch (Exception e)
            {
                throw new Exception($"获取 Workbook 出错：{e.Message}", e);
            }
        }

        /// <summary>
        /// 获取单元格合并信息
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <param name="cell">单元格</param>
        /// <returns>MergedInfo</returns>
        public virtual ExcelCellMergedInfo GetCellMergedInfo(ISheet sheet, ICell cell)
        {
            if (cell?.IsMergedCell == true)
            {
                // 得到一个sheet中有多少个合并单元格
                int sheetMergerCount = sheet.NumMergedRegions;
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    // 得出具体的合并单元格
                    CellRangeAddress ca = sheet.GetMergedRegion(i);
                    // 通过合并单元格的起始行, 结束行, 起始列, 结束列 来判断该单元格是否在合并单元格范围之内, 如果是, 则返回 true
                    if (cell.ColumnIndex <= ca.LastColumn && cell.ColumnIndex >= ca.FirstColumn)
                    {
                        if (cell.RowIndex <= ca.LastRow && cell.RowIndex >= ca.FirstRow)
                        {
                            return new ExcelCellMergedInfo(cell, true, ca);
                        }
                    }
                }
            }

            return new ExcelCellMergedInfo(cell);
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public virtual object GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            try
            {
                object value = null;
                switch (cell.CellType)
                {
                    case CellType.Blank: //空值 3
                        value = null;
                        break;
                    case CellType.Unknown: //未知 -1
                    case CellType.String: //字符串型 1
                        value = cell.StringCellValue;
                        break;
                    case CellType.Boolean: //布尔 4
                        value = cell.BooleanCellValue;
                        break;
                    case CellType.Error: //错误 5
                        try
                        {
                            value = ErrorConstant.ValueOf(cell.ErrorCellValue).Text;
                        }
                        catch
                        {
                            value = cell.ErrorCellValue;
                        }
                        break;
                    case CellType.Numeric: //数值型 0

                        if (DateUtil.IsCellDateFormatted(cell) || DateUtil.IsCellInternalDateFormatted(cell))
                        {
                            value = DateTime.FromOADate(cell.NumericCellValue);
                        }
                        else
                        {
                            value = cell.NumericCellValue;
                        }
                        break;
                    case CellType.Formula: //公式型 2
                        try
                        {
                            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                            value = GetCellValue(eva.Evaluate(cell), cell);
                        }
                        catch
                        {
                            XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                            value = GetCellValue(e.Evaluate(cell), cell);
                        }
                        break;
                    default:
                        value = cell.StringCellValue;
                        break;
                }

                return value;
            }
            catch (Exception e)
            {
                throw new Exception($"获取单元格值出错[{new CellReference(cell).FormatAsString()}]： {e.Message}", e);
            }
        }
        /// <summary>
        /// 得到公式单元格的值
        /// </summary>
        /// <param name="formulaValue"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public virtual object GetCellValue(CellValue formulaValue, ICell cell)
        {
            if (formulaValue == null || cell == null)
            {
                return formulaValue;
            }
            object value = null;
            switch (formulaValue.CellType)
            {
                case CellType.Blank:
                    value = null;
                    break;
                case CellType.Unknown:
                case CellType.String:
                    value = formulaValue.StringValue;
                    break;
                case CellType.Boolean:
                    value = formulaValue.BooleanValue.ToString(CultureInfo.CurrentCulture);
                    break;
                case CellType.Error:
                    try
                    {
                        value = ErrorConstant.ValueOf(cell.ErrorCellValue).Text;
                    }
                    catch
                    {
                        value = cell.ErrorCellValue.ToString();
                    }
                    break;
                case CellType.Numeric:
                    value = formulaValue.NumberValue.ToString(CultureInfo.CurrentCulture);
                    break;
                case CellType.Formula:
                    try
                    {
                        HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        value = GetCellValue(eva.Evaluate(cell), cell);
                    }
                    catch
                    {
                        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        value = GetCellValue(e.Evaluate(cell), cell);
                    }
                    break;
                default:
                    value = formulaValue.StringValue;
                    break;
            }

            return value;
        }

        /// <summary>
        /// 获取所在合并的单元格区域的值
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <param name="cell">单元格</param>
        /// <returns>若是合并的单元格：返回合并区域的第一个值。若是非合并单元格：返回当前表格的值</returns>
        public virtual object GetMergedCellValue(ISheet sheet, ICell cell)
        {
            var info = GetCellMergedInfo(sheet, cell);
            if (info.IsMergedRegion)
            {
                IRow fRow = sheet.GetRow(info.CellRangeAddress.FirstRow);
                ICell fCell = fRow.GetCell(info.CellRangeAddress.FirstColumn);
                return GetCellValue(fCell);
            }
            return GetCellValue(cell);
        }

        /// <summary>
        /// 获取单元格值的默认格式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public virtual short GetDefaultFormat(IWorkbook workbook, CellValueDefaultFormatEnum cellValue = CellValueDefaultFormatEnum.文本)
        {
            IDataFormat format = workbook.CreateDataFormat();
            switch (cellValue)
            {
                case CellValueDefaultFormatEnum.日期:
                    return format.GetFormat("yyyy-MM-dd HH:mm:ss");
                case CellValueDefaultFormatEnum.时间:
                    return format.GetFormat("HH:mm:ss");
                case CellValueDefaultFormatEnum.数字:
                    return format.GetFormat("0.00"); //小数点后有几个0表示精确到小数点后几位
                case CellValueDefaultFormatEnum.金额:
                    return format.GetFormat("￥#,##0.00_ ");
                case CellValueDefaultFormatEnum.百分比:
                    return format.GetFormat("0.00%");
                case CellValueDefaultFormatEnum.中文大写:
                    return format.GetFormat("[DbNum2][$-804]0");
                case CellValueDefaultFormatEnum.科学计数法:
                    return format.GetFormat("0.00E+00");
                case CellValueDefaultFormatEnum.文本:
                    return HSSFDataFormat.GetBuiltinFormat("@");
                default:
                    return HSSFDataFormat.GetBuiltinFormat("@");
            }
        }

        /// <summary>
        /// 转换列值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">当前列下标,起始0</param>
        /// <param name="valueType">值类型/属性类型，如 PropertyInfo.PropertyType ，typeof(int?)，typeof(bool),typeof(string)</param>
        /// <returns></returns>
        public virtual object ConverterCellValue(IRow row, int columnIndex, Type valueType)
        {
            ICell cell = row?.GetCell(columnIndex);
            if (cell == null)
            {
                return null;
            }

            object cellValue = GetMergedCellValue(cell.Sheet, cell);

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
        /// 获取默认的单元格样式
        /// <para>在.xls工作簿中最多可以定义4000种样式,所以该方法要在循环外定义</para>
        /// </summary>
        /// <param name="workbook">Workbook</param>
        /// <param name="style">CellStyleEnum</param>
        /// <returns></returns>
        public virtual ICellStyle GetDefaultCellStyle(IWorkbook workbook, ExcelCellStyleEnum style = ExcelCellStyleEnum.默认)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            ////边框  
            //cellStyle.BorderTop = BorderStyle.Dotted;
            //cellStyle.BorderBottom = BorderStyle.Dotted;
            //cellStyle.BorderLeft = BorderStyle.Hair;
            //cellStyle.BorderRight = BorderStyle.Hair;

            ////边框颜色  
            //cellStyle.BottomBorderColor = HSSFColor.Blue.Index;
            //cellStyle.TopBorderColor = HSSFColor.Blue.Index;

            ////背景
            //cellStyle.FillBackgroundColor = HSSFColor.Blue.Index;
            //cellStyle.FillForegroundColor = HSSFColor.Blue.Index;
            //cellStyle.FillForegroundColor = HSSFColor.White.Index;
            //cellStyle.FillPattern = FillPattern.NoFill;
            //cellStyle.FillBackgroundColor = HSSFColor.Blue.Index;

            //水平对齐  
            cellStyle.Alignment = HorizontalAlignment.Center;
            //垂直对齐  
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            //自动换行  
            cellStyle.WrapText = true;
            //缩进  
            cellStyle.Indention = 0;

            //字体
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 10;//设置字体大小 
            font.FontName = "微软雅黑";//字体名称
            font.Color = HSSFColor.Black.Index;//字体颜色

            switch (style)
            {
                case ExcelCellStyleEnum.网址:
                    font.Color = HSSFColor.Blue.Index;
                    font.IsItalic = true;
                    font.Underline = FontUnderlineType.Single;
                    break;
                case ExcelCellStyleEnum.主标题:
                    font.IsBold = true;
                    font.FontHeightInPoints = 22;
                    break;
                case ExcelCellStyleEnum.表头:
                    font.IsBold = true;
                    font.FontHeightInPoints = 12;
                    break;
                default:
                    break;
            }
            cellStyle.SetFont(font);

            return cellStyle;
        }

        /// <summary>
        /// 单列设置列宽（该方法必须在创建列后才能设置，不能在创建列前设置；列宽自动调整，必须有列数据才能处理））
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列下标（起始下表：0）</param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        /// <param name="columnAutoSize">是否自动调整</param>
        public virtual void SetColumnWidth(ISheet sheet, int columnIndex, int columnSize, bool columnAutoSize)
        {
            if (columnSize > 0)
            {
                //创建“列”后才能设置，不能在创建“列”前设置
                sheet.SetColumnWidth(columnIndex, (columnSize > 255 ? 255 : columnSize) * 256);
            }
            else if (columnAutoSize)
            {
                //列宽自动调整，必须有“列数据”才能处理）
                sheet.AutoSizeColumn(columnIndex);
            }
        }

        /// <summary>
        /// 统一设置列宽（该方法必须在创建sheet后、创建列前 设置，不能在创建列后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnSize">列尺寸（单位：字符，[0-255]）</param>
        public virtual void SetColumnWidth(ISheet sheet, int columnSize)
        {
            if (columnSize > 0)
            {
                //必须在创建“sheet”后、创建“列“前设置，不能在创建列后才设置
                sheet.DefaultColumnWidth = columnSize > 255 ? 255 : columnSize;//这里不*256
            }
        }

        /// <summary>
        /// 单行设置行高（该方法必须在创建行后才能设置，不能在创建行前设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">行</param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        public virtual void SetRowHeight(ISheet sheet, IRow row, short rowHeight)
        {
            if (rowHeight > 0)
            {
                //必须在创建“行”后才能设置，不能在创建“行”前设置
                row.Height = (short)((rowHeight > 409 ? 409 : rowHeight) * 20.0);
            }
        }

        /// <summary>
        /// 统一设置行高（该方法必须在创建sheet后马上设置，不能在创建行后才设置）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowHeight">行高（单位：磅，[0-409]））</param>
        public virtual void SetRowHeight(ISheet sheet, short rowHeight)
        {
            if (rowHeight > 0)
            {
                //必须在创建“sheet”后马上设置，不能在创建“行”后才设置
                sheet.DefaultRowHeight = (short)((rowHeight > 409 ? 409 : rowHeight) * 20.0);
            }
        }

        /// <summary>
        /// IWorkbook写入流并转换为字节
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public virtual byte[] GetAsByteArray(IWorkbook workbook)
        {
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// 合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow">起始行（起始下标：0）</param>
        /// <param name="lastRow">结束行（起始下标：0）</param>
        /// <param name="firstCol">起始列（起始下标：0）</param>
        /// <param name="lastCol">结束列（起始下标：0）</param>
        public virtual void MergedRegion(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            //合并列
            sheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }

        /// <summary>
        /// 获取区域字符串（如： "A1"）
        /// </summary>
        /// <param name="rowIndex">行下标（起始下标：0）</param>
        /// <param name="columnIndex">列下标（起始下标：0）</param>
        public virtual string GetCellAddress(int rowIndex, int columnIndex)
        {
            return new CellReference(rowIndex, columnIndex).FormatAsString();
        }

        /// <summary>
        /// 获取区域字符串（如： "A1"）
        /// </summary>
        /// <param name="cell">单元格</param>
        public virtual string GetCellAddress(ICell cell)
        {
            return new CellReference(cell).FormatAsString();
        }
    }



    /// <summary>
    /// 合并的单元格区域信息类
    /// </summary>
    public class ExcelCellMergedInfo
    {
        /// <summary>
        /// 是否合并区域
        /// </summary>
        public bool IsMergedRegion { get; set; }
        /// <summary>
        /// 单元格范围
        /// </summary>
        public CellRangeAddress CellRangeAddress { get; set; }
        /// <summary>
        /// 单元格
        /// </summary>
        public ICell ICell { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelCellMergedInfo()
        {
            IsMergedRegion = false;
            CellRangeAddress = null;
        }
        /// <summary>
        /// 构造
        /// </summary>
        public ExcelCellMergedInfo(ICell cell) : this()
        {
            ICell = cell;
        }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="isMergedRegion"></param>
        /// <param name="cellRangeAddress"></param>
        public ExcelCellMergedInfo(ICell cell, bool isMergedRegion, CellRangeAddress cellRangeAddress)
        {
            ICell = cell;
            IsMergedRegion = isMergedRegion;
            CellRangeAddress = cellRangeAddress;
        }
    }

    /// <summary>
    /// 单元格数据类型
    /// </summary>
    public enum CellValueDefaultFormatEnum
    {
        文本,
        数字,
        日期,
        时间,
        金额,
        百分比,
        中文大写,
        科学计数法,
    }

    /// <summary>
    /// 单元格样式类型
    /// </summary>
    public enum ExcelCellStyleEnum
    {
        默认,
        主标题,
        表头,
        网址,
    }

}
