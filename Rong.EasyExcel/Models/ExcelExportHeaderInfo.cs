using System.Reflection;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 导出的表头信息
    /// </summary>
    public class ExcelExportHeaderInfo
    {
        /// <summary>
        /// 对应的属性
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 显示的表头名称
        /// </summary>
        public string HeaderName { get; set; }
    }
}
