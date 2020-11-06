using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Rong.EasyExcel
{
    /// <summary>
    /// Excel 导入服务
    /// <para>请看例子 <see cref="ExcelDemo"/></para>
    /// </summary>
    public interface IExcelImportManager
    {
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
        List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(
            string filePhysicalPath,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new();

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
        Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(
              string filePhysicalPath,
              Action<ExcelImportOptions> optionAction = null
          ) where TImportDto : class, new();

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
        List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(
            byte[] fileBytes,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new();

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
        Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(
            byte[] fileBytes,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileStream">excel 文件流</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        List<ExcelSheetDataOutput<TImportDto>> Import<TImportDto>(
            Stream fileStream,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="TImportDto">表头对应的类
        /// <para>1.表头名称对应 <see cref="System.ComponentModel.DataAnnotations"/> 下的 DisplayName 特性</para>
        /// <para>2.字段验证可使用 <see cref="System.ComponentModel.DataAnnotations"/> 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType，DefaultValue 等】</para>
        /// </typeparam>
        /// <param name="fileStream">excel 文件流</param>
        /// <param name="optionAction">配置选项</param>
        /// <returns></returns>
        Task<List<ExcelSheetDataOutput<TImportDto>>> ImportAsync<TImportDto>(
            Stream fileStream,
            Action<ExcelImportOptions> optionAction = null
        ) where TImportDto : class, new();
    }
}
