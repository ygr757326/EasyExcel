using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Rong.EasyExcel
{
    /// <summary>
    /// Excel 导入服务
    /// </summary>
    public abstract class ExcelImportManager : IExcelImportManager
    {
        /// <summary>
        /// 构造
        /// </summary>
        protected ExcelImportManager()
        {
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="filePhysicalPath">excel 文件路径</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(string filePhysicalPath, Action<ExcelImportOptions> optionAction = null) where TImportDto : class, new()
        {
            try
            {
                ExcelHelper.ValidationExcel(filePhysicalPath);

                using (var stream = new FileStream(filePhysicalPath, FileMode.Open, FileAccess.Read))
                {
                    return Import<TImportDto>(stream, optionAction);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="filePhysicalPath">excel 文件路径</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(string filePhysicalPath, Action<ExcelImportOptions> optionAction = null) where TImportDto : class, new()
        {
            return Task.FromResult(Import<TImportDto>(filePhysicalPath, optionAction));
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileBytes">excel 文件字节</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(byte[] fileBytes, Action<ExcelImportOptions> optionAction = null) where TImportDto : class, new()
        {
            try
            {
                using (var stream = new MemoryStream(fileBytes))
                {
                    return Import<TImportDto>(stream, optionAction);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }

        }


        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileBytes">excel 文件字节</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(byte[] fileBytes, Action<ExcelImportOptions> optionAction = null) where TImportDto : class, new()
        {
            return Task.FromResult(Import<TImportDto>(fileBytes, optionAction));
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileStream">文件流</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(
            Stream fileStream,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new()
        {
            try
            {
                return ImplementImport<TImportDto>(fileStream, optionAction);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e);
            }

        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileStream">文件流</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        public Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(Stream fileStream, Action<ExcelImportOptions> optionAction = null) where TImportDto : class, new()
        {
            return Task.FromResult(Import<TImportDto>(fileStream, optionAction));
        }

        /// <summary>
        /// 导入实现
        /// </summary>
        /// <typeparam name="TImportDto"></typeparam>
        /// <param name="fileStream"></param>
        /// <param name="optionAction"></param>
        /// <returns></returns>
        protected abstract List<ExcelSheetDataOutput<TImportDto>> ImplementImport<TImportDto>(Stream fileStream, Action<ExcelImportOptions> optionAction) where TImportDto : class, new();

    }
}
