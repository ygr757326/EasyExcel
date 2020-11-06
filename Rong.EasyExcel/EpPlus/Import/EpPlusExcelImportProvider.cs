using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace Rong.EasyExcel.EpPlus.Import
{
    /// <summary>
    /// EpPlus excel 导入服务
    /// </summary>
    public class EpPlusExcelImportProvider : ExcelImportManager
    {
        private readonly IEpPlusExcelHandle _epPlusExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusExcelImportProvider(IEpPlusExcelHandle epPlusExcelHandle)
        {
            _epPlusExcelHandle = epPlusExcelHandle;
        }
        protected override List<ExcelSheetDataOutput<TImportDto>> ImplementImport<TImportDto>(Stream fileStream, Action<ExcelImportOptions> optionAction)
        {
            EpPlusExcelImportBase import = new EpPlusExcelImportBase(_epPlusExcelHandle);

            return import.ProcessExcelFile<TImportDto>(fileStream, optionAction);
        }
    }
}
