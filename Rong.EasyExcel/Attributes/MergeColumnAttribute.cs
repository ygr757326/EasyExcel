using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel 合并列（仅导出时用）
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class MergeColumnAttribute : Attribute
    {
        /// <summary>
        /// 属性名称集合
        /// </summary>
        public string[] PropertyNames { get; }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="propertyNames">属性名称集合</param>
        public MergeColumnAttribute(params string[] propertyNames)
        {
            PropertyNames = propertyNames;
        }
    }
}
