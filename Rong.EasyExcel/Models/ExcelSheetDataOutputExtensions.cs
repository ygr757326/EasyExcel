using System;
using System.Collections.Generic;
using System.Linq;

namespace Rong.EasyExcel.Models
{
    public static class ExcelSheetDataOutputExtensions
    {
        /// <summary>
        /// 获取错误消息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static string GetErrorMessage<T>(this ExcelImportRowInfo<T> output) where T : class, new()
        {
            if (output?.Errors == null)
            {
                return null;
            }

            return $"行编号【{output.RowNum}】存在错误：{string.Join(",", output.Errors.Select(a => a.ErrorMessage))};\r\n";
        }

        /// <summary>
        /// 检查错误并抛出异常
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static void CheckError<T>(this ExcelImportRowInfo<T> output) where T : class, new()
        {
            if (output == null)
            {
                return;
            }
            if (!output.IsValid)
            {
                throw new Exception($"{output.GetErrorMessage()}");
            }
        }

        /// <summary>
        /// 获取无效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetInvalidData<T>(this IEnumerable<ExcelImportRowInfo<T>> output) where T : class, new()
        {
            return output?.GetData(false);
        }

        /// <summary>
        /// 获取有效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetValidData<T>(this IEnumerable<ExcelImportRowInfo<T>> output) where T : class, new()
        {
            return output?.GetData(true);
        }

        /// <summary>
        /// 获取所有数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAllData<T>(this IEnumerable<ExcelImportRowInfo<T>> output) where T : class, new()
        {
            return output?.GetData(null);
        }

        /// <summary>
        /// 获取错误消息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static string GetErrorMessage<T>(this ExcelSheetDataOutput<T> output) where T : class, new()
        {
            if (output == null || output.InvalidCount <= 0)
            {
                return null;
            }
            return $"工作表【{output.SheetName}】数据错误：\r\n{string.Join(" ", output.Rows.Select(a => a.GetErrorMessage()).Where(a => a != null))}\r\n";
        }

        /// <summary>
        /// 检查错误并抛出异常
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static void CheckError<T>(this ExcelSheetDataOutput<T> output) where T : class, new()
        {
            if (output == null)
            {
                return;
            }

            if (output.InvalidCount > 0)
            {
                throw new Exception(output.GetErrorMessage());
            }
        }

        /// <summary>
        /// 获取无效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetInvalidData<T>(this ExcelSheetDataOutput<T> output) where T : class, new()
        {
            return output?.GetData(false);
        }

        /// <summary>
        /// 获取有效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetValidData<T>(this ExcelSheetDataOutput<T> output) where T : class, new()
        {
            return output?.GetData(true);
        }

        /// <summary>
        /// 获取所有数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAllData<T>(this ExcelSheetDataOutput<T> output) where T : class, new()
        {
            return output?.GetData(null);
        }


        /// <summary>
        /// 获取错误消息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static string GetErrorMessage<T>(this IEnumerable<ExcelSheetDataOutput<T>> output) where T : class, new()
        {
            if (output == null)
            {
                return null;
            }

            if (!output.Any(a => a.InvalidCount > 0))
            {
                return null;
            }

            return string.Join(" ", output.Where(a => a.InvalidCount > 0).Select(a => a.GetErrorMessage()));
        }

        /// <summary>
        /// 检查错误并抛出异常
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static void CheckError<T>(this IEnumerable<ExcelSheetDataOutput<T>> output) where T : class, new()
        {
            if (output == null)
            {
                return;
            }
            if (output.Any(a => a.InvalidCount > 0))
            {
                throw new Exception(output.GetErrorMessage());
            }
        }

        /// <summary>
        /// 获取无效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetInvalidData<T>(this IEnumerable<ExcelSheetDataOutput<T>> output) where T : class, new()
        {
            return output?.GetData(false);
        }
        /// <summary>
        /// 获取有效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetValidData<T>(this IEnumerable<ExcelSheetDataOutput<T>> output) where T : class, new()
        {
            return output?.GetData(true);
        }
        /// <summary>
        /// 获取全部数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAllData<T>(this IEnumerable<ExcelSheetDataOutput<T>> output) where T : class, new()
        {
            return output?.GetData(null);
        }

        #region 私有

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <param name="isValid">是否有效数据，null则全部</param>
        /// <returns></returns>
        public static T GetData<T>(this ExcelImportRowInfo<T> output, bool? isValid = true) where T : class, new()
        {
            if (output == null)
            {
                return null;
            }

            if (isValid == null || output.IsValid == isValid)
            {
                return output.Row;
            }

            return null;
        }

        /// <summary>
        /// 获取无效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <param name="isValid">是否有效，null全部</param>
        /// <returns></returns>
        private static IEnumerable<T> GetData<T>(this IEnumerable<ExcelImportRowInfo<T>> output, bool? isValid) where T : class, new()
        {
            return output.Select(a => a?.GetData(isValid)).Where(a => a != null).Select(a => a);
        }
        /// <summary>
        /// 获取无效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <param name="isValid">是否有效，null全部</param>
        /// <returns></returns>
        private static IEnumerable<T> GetData<T>(this ExcelSheetDataOutput<T> output, bool? isValid) where T : class, new()
        {
            return output?.Rows?.GetData(isValid);
        }
        /// <summary>
        /// 获取有效数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="output"></param>
        /// <param name="isValid">是否有效，null全部</param>
        /// <returns></returns>
        private static IEnumerable<T> GetData<T>(this IEnumerable<ExcelSheetDataOutput<T>> output, bool? isValid) where T : class, new()
        {
            return output?.SelectMany(a => a.Rows).GetData(isValid);
        }
        #endregion
    }
}
