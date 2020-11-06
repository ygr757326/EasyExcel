using Rong.EasyExcel.Models;
using System;
using System.Collections.Generic;

namespace Rong.EasyExcel.Npoi.Export
{
    /// <summary>
    /// Npoi excel 导出服务
    /// </summary>
    public class NpoiExcelExportProvider : ExcelExportManager
    {
        private readonly INpoiCellStyleHandle _npoiCellStyleHandle;
        private readonly INpoiExcelHandle _npoiExcelHandle;

        /// <summary>
        /// 构造
        /// </summary>
        public NpoiExcelExportProvider(INpoiCellStyleHandle npoiCellStyleHandle, INpoiExcelHandle npoiExcelHandle)
        {
            _npoiCellStyleHandle = npoiCellStyleHandle;
            _npoiExcelHandle = npoiExcelHandle;
        }

        protected override byte[] ImplementExport<TExportDto>(List<TExportDto> data, Action<ExcelExportOptions> optionAction, string[] onlyExportHeaderName)
        {
            NpoiExcelExportBase export = new NpoiExcelExportBase(_npoiCellStyleHandle, _npoiExcelHandle);

            return export.Export<TExportDto>(data, optionAction, onlyExportHeaderName);
        }
    }
}
