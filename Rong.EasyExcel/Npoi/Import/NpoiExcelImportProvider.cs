using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace Rong.EasyExcel.Npoi.Import
{
    /// <summary>
    /// Npoi excel 导入服务
    /// </summary>
    public class NpoiExcelImportProvider : ExcelImportManager
    {
        private readonly INpoiExcelHandle _npoiExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public NpoiExcelImportProvider(INpoiExcelHandle npoiExcelHandle)
        {
            _npoiExcelHandle = npoiExcelHandle;
        }
        protected override List<ExcelSheetDataOutput<TImportDto>> ImplementImport<TImportDto>(Stream fileStream, Action<ExcelImportOptions> optionAction)
        {
            NpoiExcelImportBase import = new NpoiExcelImportBase(_npoiExcelHandle);

            return import.ProcessExcelFile<TImportDto>(fileStream, optionAction);
        }
    }
}
