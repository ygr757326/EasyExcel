using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;
using Rong.EasyExcel.EpPlus.Export;
using Rong.EasyExcel.EpPlus.Import;

namespace Rong.EasyExcel.EpPlus
{
    public static class EpPlusExcelExtensions
    {
        /// <summary>
        /// 使用 EpPlus excel导入导出
        /// </summary>
        /// <param name="services"></param>
        public static void AddEpPlusExcel(this IServiceCollection services)
        {
            services.AddSingleton<IEpPlusCellStyleHandle, EpPlusCellStyleHandle>();
            services.AddSingleton<IEpPlusExcelHandle, EpPlusExcelHandle>();

            services.AddTransient<IExcelImportManager, EpPlusExcelImportProvider>();
            services.AddTransient<IExcelExportManager, EpPlusExcelExportProvider>();
        }
    }
}
