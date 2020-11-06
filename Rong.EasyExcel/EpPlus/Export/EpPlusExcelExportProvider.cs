using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;

namespace Rong.EasyExcel.EpPlus.Export
{
    /// <summary>
    /// EpPlus excel 导出服务
    /// </summary>
    public class EpPlusExcelExportProvider : ExcelExportManager
    {
        private readonly IEpPlusCellStyleHandle _epPlusCellStyleHandle;
        private readonly IEpPlusExcelHandle _epPlusExcelHandle;
        /// <summary>
        /// 构造
        /// </summary>
        public EpPlusExcelExportProvider(IEpPlusCellStyleHandle epPlusCellStyleHandle, IEpPlusExcelHandle epPlusExcelHandle)
        {
            _epPlusCellStyleHandle = epPlusCellStyleHandle;
            _epPlusExcelHandle = epPlusExcelHandle;
        }

        protected override byte[] ImplementExport<TExportDto>(List<TExportDto> data, Action<ExcelExportOptions> optionAction, string[] onlyExportHeaderName)
        {
            EpPlusExcelExportBase export = new EpPlusExcelExportBase(_epPlusCellStyleHandle, _epPlusExcelHandle);

            return export.Export<TExportDto>(data, optionAction, onlyExportHeaderName);
        }
    }
}
