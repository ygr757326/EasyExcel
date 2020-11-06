using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel合并行（仅导出时用）
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class MergeRowAttribute : Attribute
    {
    }
}
