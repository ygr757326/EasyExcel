using Rong.EasyExcel.Attributes;
using Rong.EasyExcel.Models;
using System.Collections.Generic;
using System.Reflection;

namespace Rong.EasyExcel
{
    /// <summary>
    /// Excel 样式特性处理帮助器
    /// </summary>
    public static class ExcelStyleHelper
    {
        /// <summary>
        /// 获取表头样式和字体集合
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <returns></returns>
        public static List<ExcelCellStyleInfo<HeaderStyleAttribute, HeaderFontAttribute>> GetHeaderStyleFontList<TExportDto>() where TExportDto : class, new()
        {
            var list = new List<ExcelCellStyleInfo<HeaderStyleAttribute, HeaderFontAttribute>>();

            foreach (var p in ExcelHelper.GetProperties<TExportDto>())
            {
                list.Add(GetHeaderStyleFont<TExportDto>(p));
            }

            return list;
        }

        /// <summary>
        /// 获取数据样式和字体集合
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <returns></returns>
        public static List<ExcelCellStyleInfo<DataStyleAttribute, DataFontAttribute>> GetDataStyleFontList<TExportDto>() where TExportDto : class, new()
        {
            var list = new List<ExcelCellStyleInfo<DataStyleAttribute, DataFontAttribute>>();

            foreach (var p in ExcelHelper.GetProperties<TExportDto>())
            {
                list.Add(GetDataStyleFont<TExportDto>(p));
            }

            return list;
        }

        /// <summary>
        /// 获取表头样式和字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static ExcelCellStyleInfo<HeaderStyleAttribute, HeaderFontAttribute> GetHeaderStyleFont<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var style = GetHeaderStyleAttr<TExportDto>(m);
            var font = GetHeaderFontAttr<TExportDto>(m);

            return new ExcelCellStyleInfo<HeaderStyleAttribute, HeaderFontAttribute>(m, style, font);
        }

        /// <summary>
        /// 获取数据样式和字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static ExcelCellStyleInfo<DataStyleAttribute, DataFontAttribute> GetDataStyleFont<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var style = GetDataStyleAttr<TExportDto>(m);
            var font = GetDataFontAttr<TExportDto>(m);

            return new ExcelCellStyleInfo<DataStyleAttribute, DataFontAttribute>(m, style, font);
        }

        /// <summary>
        /// 获取表头样式
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static HeaderStyleAttribute GetHeaderStyleAttr<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var classType = typeof(TExportDto);

            var classStyleAttr = classType.GetCustomAttribute<HeaderStyleAttribute>();
            var styleAttr =m.GetCustomAttribute<HeaderStyleAttribute>();

            return styleAttr ?? classStyleAttr ?? new HeaderStyleAttribute();
        }

        /// <summary>
        /// 获取表头字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static HeaderFontAttribute GetHeaderFontAttr<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var classType = typeof(TExportDto);

            var classFontAttr = classType.GetCustomAttribute<HeaderFontAttribute>();
            var fontAttr = m.GetCustomAttribute<HeaderFontAttribute>();

            return fontAttr ?? classFontAttr ?? new HeaderFontAttribute();
        }

        /// <summary>
        /// 获取数据样式
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static DataStyleAttribute GetDataStyleAttr<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var classType = typeof(TExportDto);

            var classStyleAttr = classType.GetCustomAttribute<DataStyleAttribute>();
            var styleAttr = m.GetCustomAttribute<DataStyleAttribute>();

            return styleAttr ?? classStyleAttr ?? new DataStyleAttribute();
        }

        /// <summary>
        /// 获取数据字体
        /// </summary>
        /// <typeparam name="TExportDto"></typeparam>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static DataFontAttribute GetDataFontAttr<TExportDto>(this MemberInfo m) where TExportDto : class, new()
        {
            var classType = typeof(TExportDto);

            var classFontAttr = classType.GetCustomAttribute<DataFontAttribute>();
            var fontAttr = m.GetCustomAttribute<DataFontAttribute>();

            return fontAttr ?? classFontAttr ?? new DataFontAttribute();
        }

        /// <summary>
        /// 是否有数据字体特性
        /// </summary>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static bool HasDataFontAttr(this MemberInfo m) => m.GetCustomAttribute<DataFontAttribute>() != null;

        /// <summary>
        /// 是否有数据样式特性
        /// </summary>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static bool HasDataStyleAttr(this MemberInfo m) => m.GetCustomAttribute<DataStyleAttribute>() != null;

        /// <summary>
        /// 是否有表头样式特性
        /// </summary>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static bool HasHeaderStyleAttr(this MemberInfo m) => m.GetCustomAttribute<HeaderStyleAttribute>() != null;

        /// <summary>
        /// 是否有表头字体特性
        /// </summary>
        /// <param name="m">typeof(类)、PropertyInfo</param>
        /// <returns></returns>
        public static bool HasHeaderFontAttr(this MemberInfo m) => m.GetCustomAttribute<HeaderFontAttribute>() != null;
    }
}
