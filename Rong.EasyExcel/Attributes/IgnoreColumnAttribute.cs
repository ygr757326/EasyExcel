using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel列字段导入/导出时忽略
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class IgnoreColumnAttribute : Attribute
    {
    }
}
