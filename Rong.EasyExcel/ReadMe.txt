# EasyExcel

>当前已集成 NPOI、EpPlus 两种方式。其他方式可自行参照项目地址中的 ReadMe.txt 文档说明快速集成。

>使用方法请看 ExcelDemo.cs。 

>导入服务： IExcelImportManager 

>导出服务： IExcelExportManager

***
### 导入支持： 

单/多个sheet导入、 DataAnnotations 数据验证、默认值、字段忽略、指定行区间导入等。

***
### 导出支持： 

样式、字体、列宽、行高、列合并、行合并、列统计、数据格式化、指定表头导出 等设置。

***
### 1.添加服务.

在 Startup.cs 中 添加服务（选其一即可，多个则取最后一个）

    * 使用 Npoi

    services.AddNpoiExcel();
  
    * 使用 EpPlus

    services.AddEpPlusExcel();

 ***
### 2.导入

> excel 中的表头行的名称必须与Display特性的name或属性名称对应，导入类中的所有非忽略属性在表头中必须存在

        //导入类
        public class ImportTest
        {
            [Display(Name = "姓名")]
            [Required(ErrorMessage = "{0}不能为空")]
            [StringLength(4, ErrorMessage = "{0}最大长度为{1}")]
            public virtual string Name { get; set; }

            [Display(Name = "手机号")]
            [RegularExpression(@"^1[3456789]\d{9}$", ErrorMessage = "{0}格式错误")]
            public virtual string Phone { get; set; }

            [Display(Name = "年龄")]
            [Required(ErrorMessage = "{0}不能为空")]
            [Range(10, 100, ErrorMessage = "{0}区间为{1}~{2}")]
            public virtual int Age { get; set; }

            [Display(Name = "成绩")]
            [Range(0, 150, ErrorMessage = "{0}区间为{1}~{2}")]
            public virtual decimal? Score { get; set; }

            [Display(Name = "日期")]
            [DefaultValue(typeof(DateTime), "2020-9-9")]
            public virtual DateTime Date { get; set; }

            [Display(Name = "学历")]
            [EnumDataType(typeof(TestEnum), ErrorMessage = "{0}值不存在")]
            public virtual TestEnum? Edu { get; set; }

            public virtual string NoDisplayName { get; set; }

            [IgnoreColumn]
            public virtual string IgnoreName { get; set; }
        }

        //导入方法
        public async Task Import(Stream stream)
        {
           try
            {
                var data = await _excelImportManager.ImportAsync<ImportTest>(stream, opt =>
                {
                   // opt.SheetIndex = 0;
                   // opt.ValidateMode = ExcelValidateModeEnum.ThrowRow;//可设置异常处理模式
                });

                //检查错误并抛出异常
                data.CheckError();

                //获取有效数据
                var valid = data.GetValidData();

                //获取无效数据
                var invalid = data.GetInvalidData();

                //获取全部数据
                var all = data.GetAllData();

                //获取错误信息，若无错误则返回null
                var error = data.GetErrorMessage();
            }
            catch (Exception e)
            { 
                //返回错误信息： e.Message
            }
        }

***
#### 说明

* 表头：表头名称对应 System.ComponentModel.DataAnnotations 下的 Display 特性的 Name ，若不存在 Display 特性 ，则使用 TImportDto 属性名称作为表头
* 验证：字段验证可使用 System.ComponentModel.DataAnnotations 的所有特性，如 Required，StringLength，Range，RegularExpression，EnumDataType等
* 默认值：若某字段为空但是需要设置默认值时，可使用 DefaultValue 特性，如 [DefaultValue(typeof(DateTime), "2020-9-9")]
* 忽略字段：可使用 ColumnIgnore 特性来忽略导入的 TImportDto 的属性字段，添加该特性的属性字段不会被单元格赋值和验证
* 配置：_excelImportManager.Import 的第二个参数 ExcelImportOptions 可进行配置

***
### 3.导出

> 若无其他特殊需求，该 ExportTest 上的所有特性，只需要 Display 即可。

        //导出类
        [HeaderStyle(ColumnAutoSize = true)]
        [RowHeight(20, 30)]
        [HeaderFont(Color = 14)]
        [DataStyle(FillPattern = (short)FillPattern.SolidForeground, FillForegroundColor = HSSFColor.LightOrange.Index)]
        [DataFont(Color = 16)]
        [MergeColumn(nameof(Edu), nameof(Name11))]
        public class ExportTest
        {
            [Display(Name = "姓名")]
            [DataStyle(WrapText = true, FillPattern = (short)FillPattern.SolidForeground, FillForegroundColor = HSSFColor.Green.Index)]
            [DataFont(Color = 10)]
            [MergeRow]
            public virtual string Name { get; set; }

            [Display(Name = "生日")]
            [DataStyle(DataFormat = "yyyy\"年\"m\"月\"d\"日\";@")]
            [DefaultValue(typeof(DateTime), "2020-9-9")]
            public virtual DateTime? Date { get; set; }

            [Display(Name = "年龄")]
            [ColumnStats((int)FunctionEnum.Avg)]
            public virtual int Age { get; set; }

            [Display(Name = "成绩")]
            [HeaderFont(Color = 15)]
            [DataStyle(DataFormat = "#,##0.00_ ")]
            [ColumnStats((int)FunctionEnum.Avg, OffsetRow = 4)]
            [ColumnStats((int)FunctionEnum.Sum)]
            public virtual decimal? Score { get; set; }

            [Display(Name = "是否及格")]
            public virtual bool IsPass => Score > 3000;

            [Display(Name = "学历")]
            [IgnoreColumn]
            public virtual TestEnum? Edu { get; set; }

            [Display(Name = "学历文本")]
            public virtual string EduText => Edu?.ToString();

            [Display(Name = "时间")]
            public virtual TimeSpan Time =>TimeSpan.FromDays(1);
        }
        
        //导出方法
 	    public async Task<byte[]> Export(List<ExportTest> list)
        { 
            try
            {
               var bytes = await _excelExportManager.ExportAsync<ExportTest>(list, opt =>
                {
                    //opt.ExcelType = ExcelTypeEnum.Xlsx;
                    //opt.SheetName = "sheet名称";
                },new []{"姓名","日期"});

                return bytes;
            }
            catch (Exception e)
            { 
                //返回错误信息： e.Message
            }
        }

        // 获取excel文件
        public FileStreamResult GetExcel(byte[] buffer)
        {
            string xlsMime = "application/vnd.ms-excel";
            string xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            Stream stream = new MemoryStream(buffer);
            stream.Seek(0, SeekOrigin.Begin);

            return new FileStreamResult(stream, xlsxMime)
            {
                FileDownloadName = "导出的Excel.xlsx"
            };
        }

***
#### 说明

* 表头：表头名称对应 System.ComponentModel.DataAnnotations 下的 Display 特性的 Name ，若不存在 Display 特性 ，则使用 TImportDto 属性名称作为表头

* 默认值：若某字段为空但是需要设置默认值时，可使用 DefaultValue 特性，如 [DefaultValue(typeof(DateTime), "2020-9-9")]

* 忽略字段：可使用 ColumnIgnore 特性来忽略导出的 TExportDto 的属性字段，添加该特性的属性字段不会被导出

* 配置：_excelExportManager.Export 的第二个参数 ExcelExportOptions 可进行配置

* 指定表头和顺序：可使用第三个参数指定筛选只需要导出的字段，不指定则按 TExportDto 字段顺序导出全部，指定则按数组顺序导出【导出的表头字段可通过_excelExportManager.GetExportHeader<ExportTest>()方法获取】
        
***
#### 样式、字体

> 现只针对整列，暂未实现单个单元格的设置

> 可在类上或属性上设置。若类上存在该特性，则所有表头列/数据列都有效；若属性上存在该特性，则优先属性“特性”，类上的“特性无效”【特性无效：是整个特性，不是特性中的参数值】

* 表头样式：HeaderStyle 特性（列宽）

* 表头字体：HeaderFont 特性

* 数据样式：DataStyle 特性

* 数据字体：DataFont 特性

* 行高：RowHeight 特性

***
#### 数据合并、列统计

* 行合并：特性 [MergeRow]，设置在类的属性上；

* 列合并：特性 [MergeColumn(nameof(属性1),nameof(属性2))]，设置在类上，可多个，但单个属性在多个特性中不能重复。

* 列统计：特性[ColumnStats((short)FunctionEnum.Min)]，设置在类的属性上，可多个；可指定展示到某属性列，可设置 行偏移、显示文本、显示单位 。单个属性有多个时，须指定向下偏移行数

* 合并的优先级：当列合并[MergeColumn]中指定 属性1/属性2，且 属性1/属性2 上存在行合并特性 [MergeRow] 时，还是优先列合并，除非导出的表头（导出说明 - 指定表头和顺序）中“不足以形成列合并”时，才优先行合并。

* 不足以形成列合并：如 指定属性1/属性2/属性3 ：只导出一个 或 多个但不相邻，则各个属性采用行合并

* 自动判断列合并：如 指定属性1/属性2/属性3 ,但只导出其中两个且相邻，则这两个也会进行列合并

* 列统计指定展示到某属性列，若指定属性列存在，则展示到指定属性列，若不存在则展示到当前属性列

***
### 4.集成其他导入导出：

> 导入请继承并实现以下类：

    * ExcelImportBase.cs【参照 NpoiExcelImportBase.cs】
    * ExcelImportManager.cs 【参照 NpoiExcelImportProvider.cs】，在 xxxxExcelImportProvider 中实例化 ExcelImportBase。如：

        protected override List<ExcelSheetDataOutput<TImportDto>> ImplementImport<TImportDto>(Stream fileStream, Action<ExcelImportOptions> optionAction)
        {
            xxxExcelImportBase import = new xxxExcelImportBase();

            return import.ProcessExcelFile<TImportDto>(fileStream, optionAction);
        }

> 导出请继承并实现以下类：

   * ExcelExportBase.cs【参照 NpoiExcelExportBase.cs】
   * ExcelExportManager.cs 【参照 NpoiExcelExportProvider.cs】，在 xxxxExcelExportProvider 中实例化 ExcelExportBase。如：

        protected override byte[] ImplementExport<TExportDto>(List<TExportDto> data, Action<ExcelExportOptions> optionAction, string[] onlyExportHeaderName)
        {
            xxxExcelExportBase export = new xxxExcelExportBase();

            return export.Export<TExportDto>(data, optionAction, onlyExportHeaderName);
        }

> 扩展：

    设置 IExcelImportManager 和 IExcelExportManager 的实现【参照 NpoiExcelExtensions.cs】 如：

        public static void AddxxxExcel(this IServiceCollection services)
        {
            ……

            services.AddTransient<IExcelImportManager, xxxExcelImportProvider>();
            services.AddTransient<IExcelExportManager, xxxExcelExportProvider>();
        }


    实现扩展后可按照 步骤 1 来设置服务


