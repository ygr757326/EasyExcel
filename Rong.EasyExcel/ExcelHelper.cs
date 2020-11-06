using Rong.EasyExcel.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Rong.EasyExcel
{
    /// <summary>
    /// excel 帮助器
    /// </summary>
    public static class ExcelHelper
    {
        public static string[] Extensions = { ".xlsx", ".xls" };

        /// <summary>
        /// 判断是否是Excel文件
        /// </summary>
        /// <param name="fileName">有文件后缀的文件名</param>
        /// <returns></returns>
        public static bool IsExcel(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            string fileExt = Path.GetExtension(fileName);
            return Extensions.Contains(fileExt, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 验证是否是 Excel文件，不是则抛出异常
        /// </summary>
        /// <param name="physicalPath">excel文件物理路径</param>
        /// <returns></returns>
        public static void ValidationExcel(string physicalPath)
        {
            if (string.IsNullOrWhiteSpace(physicalPath))
            {
                throw new ArgumentNullException(nameof(physicalPath), $"文件路径不能为空");
            }

            if (!File.Exists(physicalPath))
            {
                throw new FileNotFoundException($"文件不存在：{physicalPath}");
            }

            if (!IsExcel(physicalPath))
            {
                throw new Exception($"仅支持文件扩展名为 {string.Join(",", Extensions)} 的文件");
            }
        }

        /// <summary>
        /// 获取验证结果信息
        /// </summary>
        /// <param name="instance">对象</param>
        /// <returns>如果无错误，则返回 null </returns>
        public static List<ValidationResult> GetValidationResult(object instance)
        {
            if (instance == null) return null;

            List<ValidationResult> valid = new List<ValidationResult>();
            var success = Validator.TryValidateObject(instance, new ValidationContext(instance), valid, true);
            if (!success)
            {
                return valid;
            }
            return null;
        }

        /// <summary>
        /// 是数值类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsDouble(this Type type)
        {
            return type == typeof(decimal) || type == typeof(int) || type == typeof(float) ||
                  type == typeof(long) || type == typeof(sbyte) || type == typeof(short) ||
                  type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort)
                  || type == typeof(decimal?) || type == typeof(int?) || type == typeof(float?) ||
                  type == typeof(long?) || type == typeof(sbyte?) || type == typeof(short?) ||
                  type == typeof(uint?) || type == typeof(ulong?) || type == typeof(ushort?);
        }

        /// <summary>
        /// 是日期类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsDateTime(this Type type)
        {
            return type == typeof(DateTime) || type == typeof(DateTime?);
        }

        /// <summary>
        /// 是时间类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsTimeSpan(this Type type)
        {
            return type == typeof(TimeSpan) || type == typeof(TimeSpan?);
        }

        /// <summary>
        /// 是布尔类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsBool(this Type type)
        {
            return type == typeof(bool) || type == typeof(bool?);
        }

        /// <summary>
        /// 获取属性
        /// </summary>
        /// <typeparam name="TDto">导入或导出类</typeparam>
        public static PropertyInfo[] GetProperties<TDto>() where TDto : class, new()
        {
            var dtoType = typeof(TDto);
            var properties = dtoType.GetProperties()
                .Where(a => a.GetCustomAttribute<IgnoreColumnAttribute>() == null)
                .ToArray();
            return properties;
        }

        /// <summary>
        /// 获取属性的 Display.Name 集合
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDisplayNameListFromProperty<TDto>() where TDto : class, new()
        {
            return GetProperties<TDto>().Select(GetDisplayNameFromProperty).ToList();
        }

        /// <summary>
        /// 获取属性的 Display.Name
        /// <para>若未设置Display.Name，则返回 field.Name</para>
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public static string GetDisplayNameFromProperty(this PropertyInfo property)
        {
            if (property == null)
            {
                return null;
            }
            return property.GetCustomAttribute<DisplayAttribute>()?.Name ?? property.Name;
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T GetTypedCellValue<T>(this object value) where T : struct
        {
            if (value == null)
                return default;
            Type type1 = value.GetType();
            Type type2 = typeof(T);
            Type type3 = !type2.IsGenericType || !(type2.GetGenericTypeDefinition() == typeof(Nullable<>)) ? null : Nullable.GetUnderlyingType(type2);
            if (type1 == type2 || type1 == type3)
                return (T)value;
            if (type3 != null && type1 == typeof(string) && ((string)value).Trim() == string.Empty)
                return default;
            Type type4 = type3;
            if ((object)type4 == null)
                type4 = type2;
            Type conversionType = type4;
            if (conversionType == typeof(DateTime))
            {
                if (value is double d)
                    return (T)((ValueType)DateTime.FromOADate(d));
                if (type1 == typeof(TimeSpan))
                    return (T)(ValueType)new DateTime(((TimeSpan)value).Ticks);
                if (type1 == typeof(string))
                    return (T)(ValueType)DateTime.Parse(value.ToString());
            }
            else if (conversionType == typeof(TimeSpan))
            {
                if (value is double d)
                    return (T)(ValueType)new TimeSpan(DateTime.FromOADate(d).Ticks);
                if (type1 == typeof(DateTime))
                    return (T)(ValueType)new TimeSpan(((DateTime)value).Ticks);
                if (type1 == typeof(string))
                    return (T)(ValueType)TimeSpan.Parse(value.ToString());
            }
            return (T)Convert.ChangeType(value, conversionType);
        }
    }
}
