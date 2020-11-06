using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Rong.EasyExcel
{
    /// <summary>
    ///  excel导入基类
    /// </summary>
    /// <typeparam name="TWorkbook">工作册</typeparam>
    /// <typeparam name="TSheet">工作表</typeparam>
    /// <typeparam name="TRow">行</typeparam>
    /// <typeparam name="TCell">单元格</typeparam>
    public abstract class ExcelImportBase<TWorkbook, TSheet, TRow, TCell>
    {
        /// <summary>
        /// 构造
        /// </summary>
        protected ExcelImportBase()
        {
        }

        /// <summary>
        /// 处理excel文件
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性，字段验证也可使用其下的所有特性，如 Required，StringLength，Range，RegularExpression 等】</para>
        /// </typeparam>
        /// <param name="fileBytes">excel 文件字节</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public List<ExcelSheetDataOutput<TImportDto>> ProcessExcelFile<TImportDto>(
            byte[] fileBytes,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new()
        {
            try
            {
                using (var stream = new MemoryStream(fileBytes))
                {
                    return ProcessExcelFile<TImportDto>(stream, optionAction);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }
        }

        /// <summary>
        /// 处理excel文件
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性，字段验证也可使用其下的所有特性，如 Required，StringLength，Range，RegularExpression 等】</para>
        /// </typeparam>
        /// <param name="fileStream">文件流</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public List<ExcelSheetDataOutput<TImportDto>> ProcessExcelFile<TImportDto>(
            Stream fileStream,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new()
        {
            try
            {
                //设置、验证 配置
                ExcelImportOptions options = new ExcelImportOptions();
                optionAction?.Invoke(options);
                options.CheckError();

                return ProcessWorkbook<TImportDto>(fileStream, options);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }
        }

        #region 私有

        /// <summary>
        /// 处理excel文件
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        private List<ExcelSheetDataOutput<TImportDto>> ProcessWorkbook<TImportDto>(Stream fileStream, ExcelImportOptions options) where TImportDto : class, new()
        {
            var dataList = new List<ExcelSheetDataOutput<TImportDto>>();

            //工作册
            TWorkbook workbook = GetWorkbook(fileStream);

            //工作表总数
            int sheetsCount = GetWorksheetNumber(workbook);

            if (options.SheetIndex > sheetsCount)
            {
                throw new Exception($"工作表 sheet 编号超出：最大只能为 {sheetsCount}");
            }

            //设置工作表数据
            if (options.SheetIndex <= 0)
            {
                //全部 Sheet
                for (int i = 0; i < sheetsCount; i++)
                {
                    var data = ProcessWorksheet<TImportDto>(workbook, i, options);
                    dataList.Add(data);

                    //验证模式
                    if (data.InvalidCount > 0)
                    {
                        if (options.ValidateMode.Equals(ExcelValidateModeEnum.StopSheet))
                        {
                            break;
                        }
                        if (options.ValidateMode.Equals(ExcelValidateModeEnum.ThrowSheet))
                        {
                            data.CheckError();
                        }
                    }
                }
            }
            else
            {
                //单个 Sheet
                var data = ProcessWorksheet<TImportDto>(workbook, options.SheetIndex - 1, options);
                dataList.Add(data);

                //验证模式
                if (dataList.Any(a => a.InvalidCount > 0))
                {
                    if (options.ValidateMode.Equals(ExcelValidateModeEnum.ThrowSheet))
                    {
                        dataList.CheckError();
                    }
                }
            }

            //验证模式
            if (dataList.Any(a => a.InvalidCount > 0))
            {
                if (options.ValidateMode.Equals(ExcelValidateModeEnum.ThrowBook))
                {
                    dataList.CheckError();
                }
            }

            return dataList;
        }

        /// <summary>
        /// 获取工作表数据
        /// </summary>
        /// <typeparam name="TImportDto"></typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="sheetIndex">工作表下标（其实下表：0）</param>
        /// <param name="options">选项配置</param>
        /// <returns></returns>
        private ExcelSheetDataOutput<TImportDto> ProcessWorksheet<TImportDto>(TWorkbook workbook, int sheetIndex, ExcelImportOptions options) where TImportDto : class, new()
        {
            //获取工作表
            TSheet worksheet = GetWorksheet(workbook, sheetIndex);

            //工作表名称
            string sheetName = GetWorksheetName(workbook, worksheet);

            try
            {
                //获取表头行
                TRow headerRow = GetHeaderRow(workbook, worksheet, options);

                //获取表头单元格集合
                List<ExcelHeaderCell> headerCells = GetHeaderCells(workbook, worksheet, headerRow);

                //表头单元格信息
                ExcelHeaderCellInfo headerCellInfo = new ExcelHeaderCellInfo(sheetName, sheetIndex, headerCells);

                //验证表头行
                ValidateHeaderRow<TImportDto>(headerCellInfo);

                //获取表头单元格属性信息
                ExcelHeaderCellPropertyInfo headerCellProperties = GetHeaderCellProperties<TImportDto>(headerCellInfo);

                //设置工作表数据
                return new ExcelSheetDataOutput<TImportDto>
                {
                    SheetName = sheetName,
                    SheetIndex = sheetIndex + 1,
                    Rows = ProcessWorksheetData<TImportDto>(workbook, worksheet, headerCellInfo, headerCellProperties, options),
                };

            }
            catch (Exception e)
            {
                throw new Exception($"工作表【{sheetName}】存在以下错误：{e.Message}", e);
            }
        }
        /// <summary>
        /// 处理工作表
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headerCellInfo"></param>
        /// <param name="headerCellProperties">表头单元格信息集合</param>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        private List<ExcelImportRowInfo<TImportDto>> ProcessWorksheetData<TImportDto>(TWorkbook workbook, TSheet worksheet, ExcelHeaderCellInfo headerCellInfo, ExcelHeaderCellPropertyInfo headerCellProperties, ExcelImportOptions options) where TImportDto : class, new()
        {
            var rows = new List<ExcelImportRowInfo<TImportDto>>();

            //获取数据行区域索引
            var rowRangeIndex = GetDataRowStartAndEndRowIndex(workbook, worksheet, options);

            for (var i = rowRangeIndex.StartIndex; i <= rowRangeIndex.EndIndex; i++)
            {
                //获取数据行
                TRow row = GetDataRow(workbook, worksheet, i);

                //获取行数据
                TImportDto entity = GetRowData<TImportDto>(workbook, worksheet, headerCellProperties, row);
                if (entity != null)
                {
                    //验证数据
                    var valid = ExcelHelper.GetValidationResult(entity);

                    var rowInfo = new ExcelImportRowInfo<TImportDto>
                    {
                        SheetName = headerCellInfo.SheetName,
                        SheetIndex = headerCellInfo.SheetIndex,
                        Row = entity,
                        Errors = valid,
                        RowNum = i + 1,
                        IsValid = valid == null
                    };

                    rows.Add(rowInfo);

                    if (!rowInfo.IsValid)
                    {
                        if (options.ValidateMode.Equals(ExcelValidateModeEnum.StopRow))
                        {
                            break;
                        }

                        if (options.ValidateMode.Equals(ExcelValidateModeEnum.ThrowRow))
                        {
                            rowInfo.CheckError();
                        }
                    }
                }
            }
            return rows;
        }

        /// <summary>
        /// 获取行数据
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headerCellProperties">表头单元格信息集合</param>
        /// <param name="dataRow">数据行</param>
        /// <returns></returns>
        private TImportDto GetRowData<TImportDto>(TWorkbook workbook, TSheet worksheet, ExcelHeaderCellPropertyInfo headerCellProperties, TRow dataRow) where TImportDto : class, new()
        {
            var data = new TImportDto();

            foreach (var p in headerCellProperties.HeaderCells)
            {
                try
                {
                    PropertyInfo property = p.PropertyInfo;
                    //if (property.SetMethod == null)
                    //{
                    //    continue;
                    //}

                    //转换单元格数据
                    object value = ConvertCellValue(workbook, worksheet, dataRow, p.ColumnIndex, property);
                    if (value == null)
                    {
                        var defaultValue = property.GetCustomAttribute<DefaultValueAttribute>();
                        if (defaultValue != null)
                        {
                            value = defaultValue.Value;
                        }
                    }

                    property.SetValue(data, value, null);
                }
                catch (Exception e)
                {
                    string cellAddress = GetCellAddress(workbook, worksheet, dataRow, p.ColumnIndex);

                    throw new Exception($"表头【{ p.Name }】在【{cellAddress}】的数据转换失败：{e.Message}", e);
                }
            }
            return data;
        }

        /// <summary>
        /// 验证表头行
        /// </summary>
        /// <param name="headerCellInfo">表头单元格信息集合</param>
        private void ValidateHeaderRow<TImportDto>(ExcelHeaderCellInfo headerCellInfo) where TImportDto : class, new()
        {
            if (headerCellInfo.HeaderCells?.Any() != true)
            {
                throw new Exception($"表头行不能为空");
            }

            //属性名称
            var propertyNames = ExcelHelper.GetDisplayNameListFromProperty<TImportDto>();

            if (!propertyNames.Any())
            {
                throw new Exception($"类 {typeof(TImportDto).Name} 对应的表头不能为空");
            }
            var propertyDuplicate = propertyNames.GroupBy(a => a).Where(a => a.Count() > 1).Select(a => a.Key).ToList();
            if (propertyDuplicate.Any())
            {
                throw new Exception($"类 {typeof(TImportDto).Name} 中 Display Name 重复（或与属性名称重复）：{string.Join(",", propertyDuplicate)}");
            }

            //excel表头名称
            var headerNames = headerCellInfo.HeaderCells.Select(a => a.Name).ToList();
            var headerDuplicate = headerNames.GroupBy(a => a).Where(a => a.Count() > 1).Select(a => a.Key).ToList();
            if (headerDuplicate.Any())
            {
                throw new Exception($"表头名称重复：{string.Join(",", headerDuplicate)}");
            }

            var except = propertyNames.Except(headerNames).ToList();
            if (except.Any())
            {
                throw new Exception($"工作表中不存在以下表头名称：{string.Join(",", except)}");
            }
        }

        /// <summary>
        /// 获取表头单元格和属性
        /// </summary>
        /// <param name="headerCellInfo">表头单元格信息集合</param>
        private ExcelHeaderCellPropertyInfo GetHeaderCellProperties<TImportDto>(ExcelHeaderCellInfo headerCellInfo) where TImportDto : class, new()
        {
            if (headerCellInfo.HeaderCells?.Any() != true)
            {
                throw new Exception($"表头行不能为空");
            }

            var cellProperties = new ExcelHeaderCellPropertyInfo
            {
                SheetName = headerCellInfo.SheetName,
                SheetIndex = headerCellInfo.SheetIndex
            };

            //属性名称
            var properties = ExcelHelper.GetProperties<TImportDto>();

            foreach (var p in properties)
            {
                var name = p.GetDisplayNameFromProperty()?.Trim();
                var cell = headerCellInfo.HeaderCells.FirstOrDefault(a => a.Name.Trim() == name);
                if (cell != null)
                {
                    cellProperties.HeaderCells.Add(new ExcelHeaderCellProperty
                    {
                        Name = cell.Name,
                        ColumnIndex = cell.ColumnIndex,
                        RowIndex = cell.RowIndex,
                        PropertyInfo = p
                    });
                }
            }

            return cellProperties;
        }

        #endregion

        #region 抽象方法

        /// <summary>
        /// 获取工作册【步骤 1】
        /// </summary>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        protected abstract TWorkbook GetWorkbook(Stream fileStream);

        /// <summary>
        /// 获取工作表数量【步骤 2】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <returns></returns>
        protected abstract int GetWorksheetNumber(TWorkbook workbook);

        /// <summary>
        /// 获取工作表【步骤 3】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="sheetIndex">工作表下标（起始下标： 0）</param>
        /// <returns></returns>
        protected abstract TSheet GetWorksheet(TWorkbook workbook, int sheetIndex);

        /// <summary>
        /// 获取工作表名称【步骤 4】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <returns></returns>
        protected abstract string GetWorksheetName(TWorkbook workbook, TSheet worksheet);

        /// <summary>
        /// 获取表头行【步骤 5】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        protected abstract TRow GetHeaderRow(TWorkbook workbook, TSheet worksheet, ExcelImportOptions options);

        /// <summary>
        /// 获取表头单元格信息集合【步骤 6）
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headerRow">表头行</param>
        protected abstract List<ExcelHeaderCell> GetHeaderCells(TWorkbook workbook, TSheet worksheet, TRow headerRow);

        /// <summary>
        /// 获取数据行的 起始、结束行下标编号（起始下标：0）【步骤 7】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        protected abstract ExcelDataRowRangeIndex GetDataRowStartAndEndRowIndex(TWorkbook workbook, TSheet worksheet, ExcelImportOptions options);

        /// <summary>
        /// 获取数据行【步骤 8】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行下标（起始下标： 0）</param>
        /// <returns></returns>
        protected abstract TRow GetDataRow(TWorkbook workbook, TSheet worksheet, int rowIndex);

        /// <summary>
        /// 转换单元格数据【步骤 9】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="dataRow">数据行</param>
        /// <param name="columnIndex">列下标（起始下标：原值）</param>
        /// <param name="property">表头对应的 TImportDto 字段属性</param>
        /// <returns></returns>
        protected abstract object ConvertCellValue(TWorkbook workbook, TSheet worksheet, TRow dataRow, int columnIndex, PropertyInfo property);

        /// <summary>
        /// 获取单元格地址，如 A1【步骤 10】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="dataRow">数据行</param>
        /// <param name="columnIndex">列下标（起始下标：原值）</param>
        /// <returns></returns>
        protected abstract string GetCellAddress(TWorkbook workbook, TSheet worksheet, TRow dataRow, int columnIndex);

        #endregion
    }
}
