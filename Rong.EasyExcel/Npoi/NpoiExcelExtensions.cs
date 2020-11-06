using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;
using Rong.EasyExcel.Npoi.Export;
using Rong.EasyExcel.Npoi.Import;

namespace Rong.EasyExcel.Npoi
{
    public static class NpoiExcelExtensions
    {
        /// <summary>
        /// 使用 NPOI excel导入导出
        /// </summary>
        /// <param name="services"></param>
        public static void AddNpoiExcel(this IServiceCollection services)
        {
            services.AddSingleton<INpoiCellStyleHandle, NpoiCellStyleHandle>();
            services.AddSingleton<INpoiExcelHandle, NpoiExcelHandle>();

            services.AddTransient<IExcelImportManager, NpoiExcelImportProvider>();
            services.AddTransient<IExcelExportManager, NpoiExcelExportProvider>();
        }
    }
}
