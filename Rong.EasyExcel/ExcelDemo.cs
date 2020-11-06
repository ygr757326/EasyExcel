using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text;
using Rong.EasyExcel.Attributes;
using Rong.EasyExcel.Models;
using System.IO;
using System.Threading.Tasks;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace Rong.EasyExcel
{
    /// <summary>
    /// excel导入例子
    /// </summary>
    public class ExcelDemo
    {
        private readonly IExcelImportManager _excelImportManager;
        private readonly IExcelExportManager _excelExportManager;
        /// <summary>
        /// 构造
        /// </summary>
        public ExcelDemo(IExcelImportManager excelImportManager, IExcelExportManager excelExportManager)
        {
            _excelImportManager = excelImportManager;
            _excelExportManager = excelExportManager;
        }

        /// <summary>
        /// 导入测试
        /// </summary>
        /// <returns></returns>
        public async Task Import(Stream stream)
        {
            try
            {
                //导入
                var data = await _excelImportManager.ImportAsync<ImportTest>(stream, opt =>
                {
                    opt.SheetIndex = 0;
                    opt.ValidateMode = ExcelValidateModeEnum.ThrowRow;
                });

                //获取有效数据
                var valid = data.GetValidData();

                //获取无效数据
                var invalid = data.GetInvalidData();

                //获取全部数据
                var all = data.GetAllData();

                //获取错误信息，若无错误则返回null
                var error = data.GetErrorMessage();

                //检查错误并抛出异常
                data.CheckError();
            }
            catch (Exception e)
            {
                //返回错误信息： e.Message
            }
        }

        /// <summary>
        /// 导出测试
        /// </summary>
        public async Task<byte[]> Export()
        {
            try
            {
                List<ExportTest> list = new List<ExportTest>();
                DateTime now = DateTime.Now.Date;
                for (int i = 0; i < 11; i++)
                {
                    list.Add(new ExportTest
                    {
                        Name = "张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三张三" + new Random().Next(1, 3),
                        Name1 = "张三" + new Random().Next(1, 3),
                        Name11 = "张三" + new Random().Next(1, 3),
                        Age = new Random().Next(10, 50),
                        Score = new Random().Next(1000, 5000),
                        Edu = (TestEnum)new Random().Next(1, 4),
                        Date = now.AddDays(new Random().Next(1, 3)),
                        Date1 = now.AddDays(new Random().Next(1, 3))
                    });
                }
                var bytes = await _excelExportManager.ExportAsync<ExportTest>(list, opt =>
                    {
                        opt.SheetName = "sheet名称";
                    }, new[] { "姓名", "日期", "年龄", "成绩" }
                );

                return bytes;
            }
            catch (Exception e)
            {
                //返回错误信息： e.Message
                throw;
            }
        }

        /// <summary>
        /// 导入类
        /// </summary>
        public class ImportTest
        {
            /// <summary>
            /// 姓名
            /// </summary>
            [Display(Name = "姓名")]
            [Required(ErrorMessage = "{0}不能为空")]
            [StringLength(4, ErrorMessage = "{0}最大长度为{1}")]
            public virtual string Name { get; set; }

            /// <summary>
            /// 手机号
            /// </summary>
            [Display(Name = "手机号")]
            [RegularExpression(@"^1[3456789]\d{9}$", ErrorMessage = "{0}格式错误")]
            public virtual string Phone { get; set; }

            /// <summary>
            /// 年龄
            /// </summary>
            [Display(Name = "年龄")]
            [Required(ErrorMessage = "{0}不能为空")]
            [Range(10, 100, ErrorMessage = "{0}区间为{1}~{2}")]
            public virtual int Age { get; set; }

            /// <summary>
            /// 成绩
            /// </summary>
            [Display(Name = "成绩")]
            [Range(0, 150, ErrorMessage = "{0}区间为{1}~{2}")]
            public virtual decimal? Score { get; set; }

            /// <summary>
            /// 日期
            /// </summary>
            [Display(Name = "日期")]
            [DefaultValue(typeof(DateTime), "2020-9-9")]
            public virtual DateTime Date { get; set; }

            /// <summary>
            /// 时间
            /// </summary>
            [Display(Name = "时间")]
            [DefaultValue(typeof(TimeSpan), "100.10:20:30")]
            public virtual TimeSpan Time { get; set; }

            /// <summary>
            /// 学历
            /// </summary>
            [Display(Name = "学历")]
            [EnumDataType(typeof(TestEnum), ErrorMessage = "{0}值不存在")]
            public virtual TestEnum? Edu { get; set; }

            /// <summary>
            /// 学历文本（忽略）
            /// </summary>
            [Display(Name = "学历文本")]
            [IgnoreColumn]
            public virtual string EduText => typeof(TestEnum).GetField(Edu.ToString())?.Name;

            /// <summary>
            /// 无 DisplayName
            /// </summary>
            public virtual string NoDisplayName { get; set; }
        }

        public enum TestEnum
        {
            小学 = 1,
            中学,
            大学
        }

        /// <summary>
        /// 导出类
        /// </summary>
        [HeaderStyle(ColumnAutoSize = true)]
        [RowHeight(20, 30)]
        [HeaderFont(Color = 14)]
        [DataStyle(FillPattern = (short)FillPattern.SolidForeground, FillForegroundColor = HSSFColor.LightOrange.Index)]
        [DataFont(Color = 16)]
        [MergeColumn(nameof(Name1), nameof(Name11))]
        [MergeColumn(nameof(Date), nameof(Date1))]
        public class ExportTest
        {
            /// <summary>
            /// 姓名
            /// </summary>
            [Display(Name = "姓名")]
            [DataStyle(WrapText = true, FillPattern = (short)FillPattern.SolidForeground, FillForegroundColor = HSSFColor.Green.Index)]
            [DataFont(Color = 10)]
            [MergeRow]
            public virtual string Name { get; set; }

            /// <summary>
            /// 姓名1
            /// </summary>
            [Display(Name = "姓名1")]
            [HeaderStyle(ColumnSize = 40)]
            [DataFont(FontHeightInPoints = 15)]
            public virtual string Name1 { get; set; }

            /// <summary>
            /// 姓名2
            /// </summary>
            [Display(Name = "姓名11")]
            [DataFont(FontHeightInPoints = 18)]
            public virtual string Name11 { get; set; }

            /// <summary>
            /// 日期
            /// </summary>
            [Display(Name = "日期")]
            [DataStyle(DataFormat = "yyyy\"年\"m\"月\"d\"日\";@")]
            [ColumnStats((int)FunctionEnum.Avg)]
            public virtual DateTime? Date { get; set; }

            /// <summary>
            /// 日期2
            /// </summary>
            [Display(Name = "日期2")]
            [DefaultValue(typeof(DateTime), "2020-9-9")]
            public virtual DateTime? Date1 { get; set; }

            /// <summary>
            /// 年龄
            /// </summary>
            [Display(Name = "年龄")]
            [ColumnStats((int)FunctionEnum.Avg)]
            public virtual int Age { get; set; }

            /// <summary>
            /// 成绩
            /// </summary>
            [Display(Name = "成绩")]
            [HeaderFont(Color = 15)]
            [DataStyle(DataFormat = "#,##0.00_ ")]
            [ColumnStats((int)FunctionEnum.Avg, OffsetRow = 4)]
            [ColumnStats((int)FunctionEnum.Sum)]
            public virtual decimal? Score { get; set; }

            /// <summary>
            /// 是否及格
            /// </summary>
            [Display(Name = "是否及格")]
            public virtual bool IsLoanCar => Score > 3000;

            /// <summary>
            /// 学历
            /// </summary>
            [Display(Name = "学历")]
            [IgnoreColumn]
            public virtual TestEnum? Edu { get; set; }

            /// <summary>
            /// 学历文本
            /// </summary>
            [Display(Name = "学历文本")]
            public virtual string EduText => Edu?.ToString();

            /// <summary>
            /// 时间
            /// </summary>
            [Display(Name = "时间")]
            public virtual TimeSpan Time => TimeSpan.FromDays(1);
        }
    }
}
