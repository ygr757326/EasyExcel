using Rong.EasyExcel.Attributes;
using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Rong.EasyExcel
{
    /// <summary>
    /// 导出基类
    /// </summary>
    public abstract class ExcelExportBase<TWorkbook, TSheet, TRow, TCell, TCellStyle>
    {
        /// <summary>
        /// 构造
        /// </summary>
        protected ExcelExportBase()
        {
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="TExportDto"><paramref name="data"/> 集合中元素的类（导出的表头顺序为字段顺序）</typeparam>
        /// <param name="data">数据</param>
        /// <param name="optionAction">配置选项</param>
        /// <param name="onlyExportHeaderName">只需要导出的表头名称（指定则按 <typeparamref name="TExportDto"/> 字段顺序导出全部，不指定空则按数组顺序导出）</param>
        /// <returns></returns>
        public byte[] Export<TExportDto>(List<TExportDto> data, Action<ExcelExportOptions> optionAction, string[] onlyExportHeaderName)
            where TExportDto : class, new()
        {
            try
            {
                ExcelExportOptions options = new ExcelExportOptions();
                optionAction?.Invoke(options);
                options.CheckError();

                //获取工作册
                TWorkbook workbook = GetWorkbook(options);

                //创建工作表
                TSheet worksheet = CreateSheet(workbook, options);

                //验证表头，并获取要导出的表头信息
                ExcelExportHeaderInfo[] headers = CheckHeader<TExportDto>(onlyExportHeaderName);

                //表头行下标
                int headerRowIndex = options.HeaderRowIndex - 1;

                //先获取所有表头列的样式和字体
                var headerStyle = GetHeaderColumnStyleAndFont<TExportDto>(workbook, worksheet);

                //处理表头单元格
                ProcessHeaderCell<TExportDto>(workbook, worksheet, headers, headerRowIndex, headerStyle);

                //数据起始行下标
                int dataRowIndex = options.DataRowStartIndex - 1;

                //先获取所有数据列的样式和字体
                var dataStyle = GetDataColumnStyleAndFont<TExportDto>(workbook, worksheet);

                //处理数据单元格
                ProcessDataCell(workbook, worksheet, headers, data, dataRowIndex, dataStyle, out int footerRowIndex);

                //处理底部数据统计
                ProcessFooterStatistics<TExportDto>(workbook, worksheet, headers, dataRowIndex, footerRowIndex, headerStyle, dataStyle);

                //处理列宽【有数据才能处理自动列宽，所以必须放到最后进行处理】
                ProcessColumnWidth<TExportDto>(workbook, worksheet, headers);

                //转换并获取工作册字节
                return GetAsByteArray(workbook, worksheet);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }
        }

        #region 私有

        /// <summary>
        /// 验证表头，并获取要导出的表头信息
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="onlyExportHeaderName">只需要导出的表头名称（指定则按 <typeparamref name="TExportDto"/> 字段顺序导出全部，不指定空则按数组顺序导出）</param>
        private ExcelExportHeaderInfo[] CheckHeader<TExportDto>(string[] onlyExportHeaderName) where TExportDto : class, new()
        {
            string className = typeof(TExportDto).Name;

            List<ExcelExportHeaderInfo> headers = new List<ExcelExportHeaderInfo>();

            var properties = GetHeaderProperties<TExportDto>();

            var headerDuplicate = properties.Select(a => a.GetDisplayNameFromProperty())
                .GroupBy(a => a)
                .Where(a => a.Count() > 1)
                .Select(a => a.Key).Distinct().ToList();

            if (headerDuplicate.Any())
            {
                throw new Exception(
                    $"类【{className}】中 Display Name 重复（或与属性名称重复）：{string.Join(",", headerDuplicate)}");
            }

            if (onlyExportHeaderName == null || onlyExportHeaderName.LongLength == 0)
            {
                headers = properties.Select(a => new ExcelExportHeaderInfo
                {
                    PropertyInfo = a,
                    HeaderName = a.GetDisplayNameFromProperty()
                }).ToList();

                if (!headers.Any())
                {
                    throw new Exception($"类【{className}】中没有要导出的表头信息");
                }
            }
            else
            {
                var onlyDuplicate = onlyExportHeaderName
                    .GroupBy(a => a)
                    .Where(a => a.Count() > 1)
                    .Select(a => a.Key).Distinct().ToList();

                if (onlyDuplicate.Any())
                {
                    throw new Exception(
                        $"指定表头名称重复：{string.Join(",", onlyDuplicate)}");
                }

                foreach (string name in onlyExportHeaderName)
                {
                    var p = properties.FirstOrDefault(a => a.GetDisplayNameFromProperty() == name);
                    if (p == null)
                    {
                        throw new Exception($"类【{className}】中未找到名称为【{name}】的 Display Name 或属性名称");
                    }

                    headers.Add(new ExcelExportHeaderInfo
                    {
                        PropertyInfo = p,
                        HeaderName = p.GetDisplayNameFromProperty()
                    });
                }
            }

            return headers.ToArray();
        }

        /// <summary>
        /// 处理表头单元格
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headers">要导出的表头信息</param>
        /// <param name="headerRowIndex">表头行下标（起始下标：0）</param>
        /// <param name="headerStyle">表头样式</param>
        private void ProcessHeaderCell<TExportDto>(TWorkbook workbook, TSheet worksheet, ExcelExportHeaderInfo[] headers, int headerRowIndex, List<ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute>> headerStyle) where TExportDto : class, new()
        {
            //处理单元格 值、样式、字体、行高
            for (int i = 0; i < headers.Length; i++)
            {
                var info = headers[i];
                int columnIndex = i;
                PropertyInfo p = info.PropertyInfo;

                //创建单元格
                TCell cell = CreateCell(workbook, worksheet, headerRowIndex, columnIndex);

                //处理表头单元格值
                ProcessHeaderCellValue(workbook, worksheet, cell, p, info.HeaderName);

                //处理表头单元格样式和字体
                var cellStyleInfo = headerStyle.FirstOrDefault(a => a.PropertyInfo == p);
                SetHeaderCellStyleAndFont<TExportDto>(workbook, worksheet, cell, cellStyleInfo);
            }

            //处理表头行 行高（必须先创建行，才能处理）
            ProcessRowHeight<TExportDto>(workbook, worksheet, headerRowIndex, true);
        }

        /// <summary>
        /// 处理数据单元格
        /// </summary>
        /// <typeparam name="TExportDto"><paramref name="data"/>集合中元素的类</typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headers">要导出的表头信息</param>
        /// <param name="data">数据集合</param>
        /// <param name="rowIndex">下一行下标（起始下标： 0）</param>
        /// <param name="dataStyle">数据样式</param>
        /// <param name="nextRowIndex">下一行下标（起始下标： 0）</param>
        /// <returns>下一行下标（从0开始）</returns>
        private void ProcessDataCell<TExportDto>(TWorkbook workbook, TSheet worksheet, ExcelExportHeaderInfo[] headers, List<TExportDto> data, int rowIndex, List<ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute>> dataStyle, out int nextRowIndex)
            where TExportDto : class, new()
        {
            //可合并行区域信息
            List<ExcelExportMergedRegionInfo> rowMergedList = new List<ExcelExportMergedRegionInfo>();
            var rowMergedHeader = GetHeaderProperties<TExportDto>().Where(a => a.GetCustomAttribute<MergeRowAttribute>() != null).Select(a => a.Name).ToList();

            //可合并列区域信息
            List<ExcelExportMergedRegionInfo> columnMergedList = new List<ExcelExportMergedRegionInfo>();
            var columnMergedHeader = typeof(TExportDto).GetCustomAttributes<MergeColumnAttribute>().Select(a => a.PropertyNames.Distinct().ToArray()).Where(a => a.Length > 1).ToList();

            //验证合并特性值
            CheckMergeAttribute<TExportDto>(columnMergedHeader);

            //处理单元格 值、样式、字体、合并
            foreach (var d in data)
            {
                for (int i = 0; i < headers.Length; i++)
                {
                    var info = headers[i];
                    int columnIndex = i;
                    PropertyInfo p = info.PropertyInfo;

                    //创建单元格
                    TCell cell = CreateCell(workbook, worksheet, rowIndex, columnIndex);

                    //处理数据单元格值
                    object value = p.GetValue(d);
                    ProcessDataCellValue(workbook, worksheet, cell, p, value);

                    //处理数据单元格样式和字体
                    var cellStyleInfo = dataStyle.FirstOrDefault(a => a.PropertyInfo == p);
                    SetDataCellStyleAndFont<TExportDto>(workbook, worksheet, cell, cellStyleInfo);

                    //处理列合并
                    ProcessMergeColumn(columnMergedHeader, columnMergedList, rowIndex, columnIndex, p, value);

                    //处理行合并
                    ProcessMergeRow(rowMergedHeader, rowMergedList, rowIndex, columnIndex, p, value);
                }

                //处理数据行 行高（必须先创建行，才能处理）
                ProcessRowHeight<TExportDto>(workbook, worksheet, rowIndex, false);

                //下一行下标
                rowIndex++;
            }

            //下一行下标
            nextRowIndex = rowIndex;

            //移除不能合并的
            columnMergedList.RemoveAll(m => !m.IsCanMergedColumn());
            rowMergedList.RemoveAll(m => !m.IsCanMergedRow());

            //若该属性存在列合并，则移除所有行合（优先列合并）
            rowMergedList.RemoveAll(m => columnMergedList.Any(a => a.PropertyNames.Intersect(m.PropertyNames).Any()));

            //所有合并信息
            var mergedRegion = rowMergedList.Concat(columnMergedList).ToList();

            //处理数据合并区域
            foreach (var m in mergedRegion)
            {
                SetMergedRegion(workbook, worksheet, m.FromRowIndex, m.ToRowIndex, m.FromColumnIndex, m.ToColumnIndex);
            }
        }

        /// <summary>
        /// 验证合并特性值
        /// </summary>
        /// <typeparam name="TExportDto">集合中元素的类</typeparam>
        /// <param name="columnMergedHeader">列合并表头信息</param>
        /// <returns>下一行下标（从0开始）</returns>
        private void CheckMergeAttribute<TExportDto>(IReadOnlyList<string[]> columnMergedHeader) where TExportDto : class, new()
        {
            string className = typeof(TExportDto).Name;
            var properties = GetHeaderProperties<TExportDto>();
            var attrName = typeof(MergeColumnAttribute).Name;

            //验证：不存在属性名称，单个属性在多个特性中重复，同一特性中属性类型不一致
            for (var index = 0; index < columnMergedHeader.Count; index++)
            {
                var names = columnMergedHeader[index];
                var num = index + 1;

                //不存在属性名称
                var noExist = names.Where(a => properties.All(p => p.Name != a)).Select(a => a);
                if (noExist.Any())
                {
                    throw new Exception(
                        $"类【{className}】的第 {num} 个 {attrName} 指定的属性名称未找到：{string.Join(",", noExist)}");
                }

                //同一特性中属性类型不一致
                var type = names.Select(n => properties.First(b => n == b.Name).PropertyType);
                if (type.Distinct().Count() > 1)
                {
                    throw new Exception(
                        $"类【{className}】的第 {num} 个 {attrName} 指定的属性类型不一致");
                }
            }

            //单个属性在多个特性中重复
            var duplicate = columnMergedHeader.SelectMany(a => a).GroupBy(a => a)
                .Where(a => a.Count() > 1)
                .Select(a => a.Key).ToList();
            if (duplicate.Any())
            {
                throw new Exception(
                    $"类【{className}】中多个合并列的属性重复：{string.Join(",", duplicate)}");
            }
        }

        /// <summary>
        /// 处理行合并
        /// </summary>
        /// <param name="rowMergedHeader">行合并表头</param>
        /// <param name="mergedList">合并信息集合</param>
        /// <param name="rowIndex">当前行下标（起始下标：0）</param>
        /// <param name="columnIndex">当前列下标（起始下标：0）</param>
        /// <param name="propertyInfo">字段属性</param>
        /// <param name="value">值</param>
        private void ProcessMergeRow(IReadOnlyList<string> rowMergedHeader, ICollection<ExcelExportMergedRegionInfo> mergedList, int rowIndex, int columnIndex, PropertyInfo propertyInfo, object value)
        {
            if (rowMergedHeader.All(a => a != propertyInfo.Name))
            {
                return;
            }

            //获取最后一个当前列的合并信息
            ExcelExportMergedRegionInfo merge = mergedList.LastOrDefault(a => a.PropertyNames.Contains(propertyInfo.Name));

            //值是否相等
            bool isValueEqual = merge?.IsValueEqual(value) == true;

            //无该列合并信息、不是同一列，值不相等但可合并、值相等但不是相邻行 都要新建合并信息
            if (merge == null || !merge.IsSameColumn(columnIndex) || !isValueEqual && merge.IsCanMergedRow() || isValueEqual && !merge.IsSiblingRow(rowIndex))
            {
                mergedList.Add(new ExcelExportMergedRegionInfo
                {
                    PropertyNames = new[] { propertyInfo.Name },
                    Value = value,
                    FromRowIndex = rowIndex,
                    ToRowIndex = rowIndex,
                    FromColumnIndex = columnIndex,
                    ToColumnIndex = columnIndex
                });
            }
            else if (!isValueEqual) //不相等，则替换掉
            {
                merge.Value = value;
                merge.FromRowIndex = rowIndex;
                merge.ToRowIndex = rowIndex;
                merge.FromColumnIndex = columnIndex;
                merge.ToColumnIndex = columnIndex;
            }
            else //值相等，相邻行 ，则改变合并行下标
            {
                if (merge.IsOutRangeRowFrom(rowIndex))
                {
                    merge.FromRowIndex -= 1;
                }
                else if (merge.IsOutRangeRowTo(rowIndex))
                {
                    merge.ToRowIndex += 1;
                }
            }
        }

        /// <summary>
        /// 处理列合并
        /// </summary>
        /// <param name="columnMergedHeader">列合并表头</param>
        /// <param name="mergedList">合并信息集合</param>
        /// <param name="rowIndex">当前行下表（起始下标：0）</param>
        /// <param name="columnIndex">当前列下标（起始下标：0）</param>
        /// <param name="propertyInfo">字段属性</param>
        /// <param name="value">值</param>
        private void ProcessMergeColumn(IReadOnlyList<string[]> columnMergedHeader, ICollection<ExcelExportMergedRegionInfo> mergedList, int rowIndex, int columnIndex, PropertyInfo propertyInfo, object value)
        {
            //处理列合并
            if (columnMergedHeader == null || !columnMergedHeader.Any(a => a.Contains(propertyInfo.Name)))
            {
                return;
            }

            //获取最后一个当前列的合并信息
            ExcelExportMergedRegionInfo merge = mergedList.LastOrDefault(a => a.PropertyNames.Contains(propertyInfo.Name));

            //值是否相等
            bool isValueEqual = merge?.IsValueEqual(value) == true;

            //无该列合并信息、不是同一行、值不相等但可合并、值相等但不是相邻列 都要新建合并信息
            if (merge == null || !merge.IsSameRow(rowIndex) || !isValueEqual && merge.IsCanMergedColumn() || isValueEqual && !merge.IsSiblingColumn(columnIndex))
            {
                mergedList.Add(new ExcelExportMergedRegionInfo
                {
                    PropertyNames = columnMergedHeader.First(a => a.Contains(propertyInfo.Name)),
                    Value = value,
                    FromRowIndex = rowIndex,
                    ToRowIndex = rowIndex,
                    FromColumnIndex = columnIndex,
                    ToColumnIndex = columnIndex
                });
            }
            else if (!isValueEqual) //值不相等，则替换掉
            {
                merge.Value = value;
                merge.FromRowIndex = rowIndex;
                merge.ToRowIndex = rowIndex;
                merge.FromColumnIndex = columnIndex;
                merge.ToColumnIndex = columnIndex;
            }
            else  //值相等，相邻列 ，则改变合并列下标
            {
                if (merge.IsOutRangeColumnFrom(columnIndex))
                {
                    merge.FromColumnIndex -= 1;
                }
                else if (merge.IsOutRangeColumnTo(columnIndex))
                {
                    merge.ToColumnIndex += 1;
                }
            }
        }

        /// <summary>
        /// 处理列宽（必须先创建列，才能处理；列宽自动调整，必须有列数据才能处理）
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headers">要导出的表头信息</param>

        private void ProcessColumnWidth<TExportDto>(TWorkbook workbook, TSheet worksheet, ExcelExportHeaderInfo[] headers) where TExportDto : class, new()
        {
            for (int i = 0; i < headers.Length; i++)
            {
                var info = headers[i];
                int columnIndex = i;
                PropertyInfo p = info.PropertyInfo;

                //设置列宽
                var styleAttr = p.GetHeaderStyleAttr<TExportDto>();
                SetColumnWidth(workbook, worksheet, columnIndex, styleAttr.ColumnSize, styleAttr.ColumnAutoSize);
            }
        }

        /// <summary>
        /// 处理行高（必须先创建行，才能处理）
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行下表</param>
        /// <param name="isHeader">是否表头</param>

        private void ProcessRowHeight<TExportDto>(TWorkbook workbook, TSheet worksheet, int rowIndex, bool isHeader) where TExportDto : class, new()
        {
            var attr = typeof(TExportDto).GetCustomAttribute<RowHeightAttribute>() ?? new RowHeightAttribute();

            SetRowHeight(workbook, worksheet, rowIndex, isHeader ? attr.HeaderRowHeight : attr.DataRowHeight);
        }

        /// <summary>
        /// 处理表头单元格值
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="propertyInfo">当前正在处理的字段属性</param>
        /// <param name="value">字段值</param>
        private void ProcessHeaderCellValue(TWorkbook workbook, TSheet worksheet, TCell cell, PropertyInfo propertyInfo, object value)
        {
            try
            {
                //设置单元格值
                SetCellValue(workbook, worksheet, cell, typeof(string), value);

            }
            catch (Exception e)
            {
                throw new Exception($"【{propertyInfo.Name}】的表头值设置出错：{e.Message}", e);
            }

        }

        /// <summary>
        /// 处理数据单元格值
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="propertyInfo">当前正在处理的字段属性</param>
        /// <param name="value">字段值</param>
        private void ProcessDataCellValue(TWorkbook workbook, TSheet worksheet, TCell cell, PropertyInfo propertyInfo, object value)
        {
            try
            {
                if (value == null)
                {
                    var defaultValue = propertyInfo.GetCustomAttribute<DefaultValueAttribute>();
                    if (defaultValue != null)
                    {
                        value = defaultValue.Value;
                    }
                }

                if (value != null)
                {
                    //设置单元格值
                    SetCellValue(workbook, worksheet, cell, propertyInfo.PropertyType, value);
                }

            }
            catch (Exception e)
            {
                throw new Exception($"【{propertyInfo.Name}】的数据值设置出错：{e.Message}", e);
            }

        }

        /// <summary>
        /// 获取所有表头列的样式和字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        private List<ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute>> GetHeaderColumnStyleAndFont<TExportDto>(TWorkbook workbook,
            TSheet worksheet) where TExportDto : class, new()
        {
            var styles = new List<ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute>>();

            //表头默认样式
            var defaultAttr = typeof(TExportDto).GetHeaderStyleFont<TExportDto>();
            TCellStyle defaultStyle =
                CreateHeaderStyleAndFont<TExportDto>(workbook, worksheet, defaultAttr.StyleAttr, defaultAttr.FontAttr);

            var properties = GetHeaderProperties<TExportDto>();
            foreach (var p in properties)
            {
                //表头样式
                TCellStyle cellStyle = defaultStyle;

                var headerAttr = p.GetHeaderStyleFont<TExportDto>();

                //设置默认格式化
                if (CanSetDefaultFormat<TExportDto>(p))
                {
                    headerAttr.StyleAttr.DataFormat = SetDefaultDataFormat(typeof(string));
                }

                //属性上有样式、有字体样式，则重新创建样式
                if (p.HasHeaderStyleAttr() || p.HasHeaderFontAttr())
                {
                    cellStyle = CreateHeaderStyleAndFont<TExportDto>(workbook, worksheet, headerAttr.StyleAttr, headerAttr.FontAttr);
                }

                //添加
                styles.Add(new ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute>(p, cellStyle, headerAttr.StyleAttr, headerAttr.FontAttr));

            }

            return styles;
        }

        /// <summary>
        /// 获取所有数据列的样式和字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        private List<ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute>> GetDataColumnStyleAndFont<TExportDto>(TWorkbook workbook,
            TSheet worksheet) where TExportDto : class, new()
        {
            var styles = new List<ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute>>();

            //数据默认样式
            var dataDefaultAttr = typeof(TExportDto).GetDataStyleFont<TExportDto>();
            TCellStyle defaultDataStyle = CreateDataStyleAndFont<TExportDto>(workbook, worksheet, dataDefaultAttr.StyleAttr,
                dataDefaultAttr.FontAttr);

            var properties = GetHeaderProperties<TExportDto>();
            foreach (var p in properties)
            {
                //数据样式
                TCellStyle cellStyle = defaultDataStyle;

                var dataAttr = p.GetDataStyleFont<TExportDto>();

                //设置默认格式化
                if (CanSetDefaultFormat<TExportDto>(p))
                {
                    dataAttr.StyleAttr.DataFormat = SetDefaultDataFormat(p.PropertyType);
                }

                //属性上有样式、有字体样式、属性为时间 则重新创建样式
                if (p.HasDataStyleAttr() || p.HasDataFontAttr() || p.PropertyType.IsDateTime())
                {
                    cellStyle = CreateDataStyleAndFont<TExportDto>(workbook, worksheet, dataAttr.StyleAttr,
                        dataAttr.FontAttr);
                }

                //添加
                styles.Add(new ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute>(p, cellStyle, dataAttr.StyleAttr, dataAttr.FontAttr));
            }

            return styles;
        }

        /// <summary>
        /// 设置默认数据格式化
        /// </summary>
        /// <param name="type">数据类型</param>
        /// <returns></returns>
        private string SetDefaultDataFormat(Type type)
        {
            if (type.IsDateTime())
            {
                return "yyyy-MM-dd";
            }

            return null;
        }

        /// <summary>
        /// 是否可以设置默认数据格式
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        private bool CanSetDefaultFormat<TExportDto>(PropertyInfo propertyInfo) where TExportDto : class, new()
        {
            var dataDefaultAttr = typeof(TExportDto).GetDataStyleFont<TExportDto>();

            var style = propertyInfo.GetDataStyleAttr<TExportDto>();

            //设置默认格式化：1.属性上有样式，但格式化为空，2.属性上没有样式，类上格式化为空
            if (propertyInfo.HasDataStyleAttr() && string.IsNullOrWhiteSpace(style.DataFormat) ||
                !propertyInfo.HasDataStyleAttr() && string.IsNullOrWhiteSpace(dataDefaultAttr.StyleAttr.DataFormat))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 获取表头属性
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <returns></returns>
        private PropertyInfo[] GetHeaderProperties<TExportDto>() where TExportDto : class, new()
        {
            return ExcelHelper.GetProperties<TExportDto>();
        }

        /// <summary>
        /// 处理底部数据统计
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="headers">要导出的表头信息</param>
        /// <param name="dataStartRowIndex">数据起始行下标（起始下标：0）</param>
        /// <param name="nextRowIndex">下一行下标（起始下标：0）</param>
        /// <param name="headerStyle">表头样式</param>
        /// <param name="dataStyle">数据样式</param>

        private void ProcessFooterStatistics<TExportDto>(TWorkbook workbook, TSheet worksheet, ExcelExportHeaderInfo[] headers, int dataStartRowIndex, int nextRowIndex, List<ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute>> headerStyle, List<ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute>> dataStyle) where TExportDto : class, new()
        {
            int dataEndRowIndex = nextRowIndex - 1;

            var properties = GetHeaderProperties<TExportDto>().Select(a => a.Name).ToList();
            var pNames = headers.Select(a => a.PropertyInfo.Name).ToList();

            for (int i = 0; i < headers.Length; i++)
            {
                var info = headers[i];
                int columnIndex = i;
                PropertyInfo p = info.PropertyInfo;

                //公式
                var fxAttrs = p.GetCustomAttributes<ColumnStatsAttribute>();
                foreach (var fxAttr in fxAttrs)
                {
                    if (!string.IsNullOrWhiteSpace(fxAttr.ShowOnColumnPropertyName) && !properties.Contains(fxAttr.ShowOnColumnPropertyName))
                    {
                        throw new Exception($"特性【{typeof(ColumnStatsAttribute).Name}】上指定的属性【{fxAttr.ShowOnColumnPropertyName}】在类【{typeof(TExportDto).Name}】中未找到");
                    }

                    var pIndex = pNames.IndexOf(fxAttr.ShowOnColumnPropertyName);
                    int pRowIndex = nextRowIndex + fxAttr.OffsetRow;
                    var pColumnIndex = pIndex == -1 ? columnIndex : pIndex;
                    var pHeaderStyle = headerStyle.FirstOrDefault(a => a.PropertyInfo == p);
                    var pDataStyle = dataStyle.FirstOrDefault(a => a.PropertyInfo == p);
                    var func = (FunctionEnum)fxAttr.Function;

                    if (fxAttr.IsShowLabel)
                    {
                        //获取标签文本
                        if (fxAttr.Label == null)
                        {
                            fxAttr.Label = $"{info.HeaderName} {typeof(FunctionEnum).GetField(func.ToString()).GetCustomAttribute<DisplayAttribute>()?.Name}";
                        }
                        if (!string.IsNullOrWhiteSpace(fxAttr.Unit))
                        {
                            fxAttr.Label += $"（{fxAttr.Unit}）";
                        }

                        //处理标签文本

                        TCell textCell = CreateCell(workbook, worksheet, pRowIndex, pColumnIndex);

                        SetCellValue(workbook, worksheet, textCell, typeof(string), fxAttr.Label);

                        //处理标签文本单元格样式和字体（采用表头样式）
                        SetHeaderCellStyleAndFont<TExportDto>(workbook, worksheet, textCell, pHeaderStyle);

                        pRowIndex++;
                    }

                    if (func.Equals(FunctionEnum.None))
                    {
                        continue;
                    }

                    //处理公式值
                    TCell cell = CreateCell(workbook, worksheet, pRowIndex, pColumnIndex);

                    //设置公式
                    string formula = GetCellFormula(workbook, worksheet, func, dataStartRowIndex, dataEndRowIndex, columnIndex, columnIndex);
                    SetCellFormula(workbook, worksheet, cell, formula);

                    //处理统计单元格样式和字体（采用数据样式）
                    SetDataCellStyleAndFont<TExportDto>(workbook, worksheet, cell, pDataStyle);
                }
            }
        }

        /// <summary>
        /// 获取公式字符串
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheet"></param>
        /// <param name="functionEnum"></param>
        /// <param name="fromRowIndex"></param>
        /// <param name="toRowIndex"></param>
        /// <param name="fromColumnIndex"></param>
        /// <param name="toColumnIndex"></param>
        /// <returns></returns>
        private string GetCellFormula(TWorkbook workbook, TSheet worksheet, FunctionEnum functionEnum, int fromRowIndex, int toRowIndex, int fromColumnIndex, int toColumnIndex)
        {
            string startAddress = GetCellAddress(workbook, worksheet, fromRowIndex, fromColumnIndex);
            string endAddress = GetCellAddress(workbook, worksheet, toRowIndex, toColumnIndex);

            string formula = null;

            switch (functionEnum)
            {
                case FunctionEnum.None:
                    {
                        formula = null;
                        break;
                    }
                case FunctionEnum.Sum:
                    {
                        formula = $"SUM({startAddress}:{endAddress})";
                        break;
                    }
                case FunctionEnum.Avg:
                    {
                        formula = $"AVERAGE({startAddress}:{endAddress})";
                        break;
                    }
                case FunctionEnum.Count:
                    {
                        formula = $"COUNT({startAddress}:{endAddress})";
                        break;
                    }
                case FunctionEnum.Max:
                    {
                        formula = $"MAX({startAddress}:{endAddress})";
                        break;
                    }
                case FunctionEnum.Min:
                    {
                        formula = $"MIN({startAddress}:{endAddress})";
                        break;
                    }
                default:
                    throw new Exception($"函数类型值【{functionEnum}】还未设置公式");
            }

            return formula;
        }

        #endregion


        #region 抽象方法

        /// <summary>
        /// 获取工作册【步骤 1】
        /// </summary>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        protected abstract TWorkbook GetWorkbook(ExcelExportOptions options);

        /// <summary>
        /// 创建工作表【步骤 2】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="options">配置选项</param>
        /// <returns></returns>
        protected abstract TSheet CreateSheet(TWorkbook workbook, ExcelExportOptions options);

        /// <summary>
        /// 创建单元格【步骤 3】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行下标（起始下标： 0）</param>
        /// <param name="columnIndex">列下标（起始下标： 0）</param>
        /// <returns></returns>
        protected abstract TCell CreateCell(TWorkbook workbook, TSheet worksheet, int rowIndex, int columnIndex);

        /// <summary>
        /// 设置数据单元格值【步骤 4】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="valueType">单元格的值类型</param>
        /// <param name="value">单元格值</param>
        protected abstract void SetCellValue(TWorkbook workbook, TSheet worksheet, TCell cell, Type valueType,
            object value);

        /// <summary>
        /// 创建表头样式和字体【步骤 5】
        /// </summary>
        /// <typeparam name="TExportDto">集合中元素的类</typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="styleAttr">样式特征</param>
        /// <param name="fontAttr">字体特征</param>
        /// <returns></returns>
        protected abstract TCellStyle CreateHeaderStyleAndFont<TExportDto>(TWorkbook workbook, TSheet worksheet,
            HeaderStyleAttribute styleAttr, HeaderFontAttribute fontAttr) where TExportDto : class, new();

        /// <summary>
        /// 创建数据样式和字体【步骤 6】
        /// </summary>
        /// <typeparam name="TExportDto">集合中元素的类</typeparam>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="styleAttr">样式特征</param>
        /// <param name="fontAttr">字体特征</param>
        /// <returns></returns>
        protected abstract TCellStyle CreateDataStyleAndFont<TExportDto>(TWorkbook workbook, TSheet worksheet,
            DataStyleAttribute styleAttr, DataFontAttribute fontAttr) where TExportDto : class, new();

        /// <summary>
        /// 设置表头单元格样式和字体【步骤 7】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="cellStyleInfo">单元格样式信息</param>
        protected abstract void SetHeaderCellStyleAndFont<TExportDto>(TWorkbook workbook, TSheet worksheet, TCell cell, ExcelCellStyleOutput<TCellStyle, HeaderStyleAttribute, HeaderFontAttribute> cellStyleInfo) where TExportDto : class, new();

        /// <summary>
        /// 设置数据单元格样式和字体【步骤 8】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="cellStyleInfo">单元格样式信息</param>
        protected abstract void SetDataCellStyleAndFont<TExportDto>(TWorkbook workbook, TSheet worksheet, TCell cell, ExcelCellStyleOutput<TCellStyle, DataStyleAttribute, DataFontAttribute> cellStyleInfo) where TExportDto : class, new();

        /// <summary>
        /// 设置列宽【步骤 9】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="columnIndex">列下标（起始下标： 0）</param>
        /// <param name="columnSize">宽度（单位：字符，取值区间：[0-255]）</param>
        /// <param name="columnAutoSize">是否自动调整</param>
        protected abstract void SetColumnWidth(TWorkbook workbook, TSheet worksheet, int columnIndex, int columnSize,
            bool columnAutoSize);

        /// <summary>
        /// 设置行高【步骤 10】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行下标（起始下标： 0）</param>
        /// <param name="rowHeight">行高（单位：磅，取值区间：[0-409]）</param>
        protected abstract void SetRowHeight(TWorkbook workbook, TSheet worksheet, int rowIndex, short rowHeight);

        /// <summary>
        /// 设置合并区域【步骤 11】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="fromRowIndex">起始行下标（起始下标： 0）</param>
        /// <param name="toRowIndex">结束行下标（起始下标： 0）</param>
        /// <param name="fromColumnIndex">起始列下标（起始下标： 0）</param>
        /// <param name="toColumnIndex">结束列下标（起始下标： 0）</param>
        protected abstract void SetMergedRegion(TWorkbook workbook, TSheet worksheet, int fromRowIndex, int toRowIndex, int fromColumnIndex, int toColumnIndex);

        /// <summary>
        /// 获取单元格地址文本（如：A1）【步骤 12】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行下标（起始下标： 0）</param>
        /// <param name="columnIndex">列下标（起始下标： 0）</param>
        /// <returns></returns>
        protected abstract string GetCellAddress(TWorkbook workbook, TSheet worksheet, int rowIndex, int columnIndex);

        /// <summary>
        /// 设置单元格公式（统计）【步骤 13】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <param name="cell">单元格</param>
        /// <param name="cellFormula">单元格公式字符串</param>
        protected abstract void SetCellFormula(TWorkbook workbook, TSheet worksheet, TCell cell, string cellFormula);

        /// <summary>
        /// 把处理好的工作册转换为字节【步骤 13】
        /// </summary>
        /// <param name="workbook">工作册</param>
        /// <param name="worksheet">工作表</param>
        /// <returns></returns>
        protected abstract byte[] GetAsByteArray(TWorkbook workbook, TSheet worksheet);

        #endregion

    }
}
