namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 导出的表头输出
    /// </summary>
    public class ExcelExportHeaderOutput
    {
        /// <summary>
        /// 显示的表头名称
        /// </summary>
        public string HeaderName { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelExportHeaderOutput()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelExportHeaderOutput(string headerName)
        {
            HeaderName = headerName;
        }
    }
}
