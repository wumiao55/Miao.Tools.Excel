using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using Miao.Tools.Excel.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace Miao.Tools.Excel
{
    /// <summary>
    /// EPPlus扩展类
    /// </summary>
    public static class EPPlusExtension
    {
        #region  写入数据

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="data">数据</param>
        public static void WriteData<T>(this ExcelWorksheet worksheet, T data) where T : class
        {
            WriteData(worksheet, data, new EPPlusConfig());
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <typeparam name="T">worksheet</typeparam>
        /// <param name="worksheet"></param>
        /// <param name="data">数据</param>
        /// <param name="config">配置信息</param>
        public static void WriteData<T>(this ExcelWorksheet worksheet, T data, EPPlusConfig config) where T : class
        {
            var start = worksheet.Dimension?.Start;
            var end = worksheet.Dimension?.End;
            if(start == null || end == null)
            {
                return;
            }

            var propertyRegex = new Regex(config.PropertyMatchRegex);
            var excelRange = worksheet.Cells[start.Row, start.Column, end.Row, end.Column];
            var excelCells = excelRange.Where(e => e.Value != null && !string.IsNullOrWhiteSpace(e.Value.ToString()) && propertyRegex.IsMatch(e.Value.ToString())).ToList();
            if (excelCells == null || excelCells.Count <= 0)
            {
                return;
            }

            var propertyInfos = data.GetType().GetProperties();
            foreach (var cell in excelCells)
            {
                object coverCellValue = null;
                var cellValue = cell.Value.ToString().Trim();
                var match = propertyRegex.Match(cellValue);
                if(match.Success && match.Groups.Count >= 2)
                {
                    string propertyName = match.Groups[1].Value.Trim();
                    var property = propertyInfos.FirstOrDefault(p => p.Name == propertyName);
                    if (property != null)
                    {
                        coverCellValue = property.GetValue(data);
                        var address = ExcelAddressUtil.ConvertExcelRowColumn(cell.Address);
                        WriteToCelll(worksheet, address.Row, address.Column, coverCellValue, config);
                    }
                }
                cell.Value = coverCellValue ?? string.Empty;
            }
        }

        #endregion

        #region 写入数据集合

        /// <summary>
        /// 写入数据集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="datas">数据集合</param>
        /// <param name="startRowNum">开始写入的行</param>
        public static void WriteDatas<T>(this ExcelWorksheet worksheet, IEnumerable<T> datas, int startRowNum) where T : class
        {
            WriteDatas(worksheet, datas, startRowNum, new EPPlusConfig());
        }

        /// <summary>
        /// 写入数据集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="datas">数据集合</param>
        /// <param name="startRowNum">开始写入的行</param>
        /// <param name="config">配置信息</param>
        public static void WriteDatas<T>(this ExcelWorksheet worksheet, IEnumerable<T> datas, int startRowNum, EPPlusConfig config) where T : class
        {
            if (datas == null || !datas.Any())
            {
                return;
            }
            var fieldDic = GetTemplateFieldDic(worksheet, config);
            var propertyInfos = datas.FirstOrDefault().GetType().GetProperties();
            Parallel.ForEach(fieldDic.Keys, columnAt =>
            {
                var fieldName = fieldDic[columnAt];
                for (int i = 0; i < datas.Count(); i++)
                {
                    var propertyInfo = propertyInfos.FirstOrDefault(p => p.Name == fieldName);
                    if (propertyInfo != null)
                    {
                        object value = propertyInfo.GetValue(datas.ElementAt(i));
                        WriteToCelll(worksheet, startRowNum + i, columnAt, value, config);
                    }
                }
            });
        }

        /// <summary>
        /// 自动填充worksheet数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="datas">数据集合</param>
        public static void AutoFill<T>(this ExcelWorksheet worksheet, IEnumerable<T> datas) where T : class
        {
            AutoFillData(worksheet, null, datas);
        }

        /// <summary>
        /// 自动填充worksheet数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="title">标题</param>
        /// <param name="datas">数据集合</param>
        public static void AutoFill<T>(this ExcelWorksheet worksheet, string title, IEnumerable<T> datas) where T : class
        {
            AutoFillData(worksheet, title, datas);
        }

        #endregion

        #region 插入图片

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="worksheet">worksheet</param>
        /// <param name="imageBytes">图片字节数据</param>
        /// <param name="rowNum">行号,从1开始</param>
        /// <param name="columnNum">列号,从1开始</param>
        /// <param name="autofit">是否自适应单元格</param>
        /// <param name="config">配置信息</param>
        public static void InsertImage(this ExcelWorksheet worksheet, byte[] imageBytes, int rowNum, int columnNum, bool autofit, EPPlusConfig config)
        {
            if (imageBytes == null || imageBytes.Length <= 0)
            {
                return;
            }
            using var image = Image.FromStream(new MemoryStream(imageBytes));
            var picture = worksheet.Drawings.AddPicture($"image_{DateTime.Now.Ticks}", image);
            var cell = worksheet.Cells[rowNum, columnNum];
            int cellColumnWidthInPix = PixelUtil.GetWidthInPixels(cell, config);
            int cellRowHeightInPix = PixelUtil.GetHeightInPixels(cell, config);
            int adjustImageWidthInPix = cellColumnWidthInPix;
            int adjustImageHeightInPix = cellRowHeightInPix;
            if (autofit)
            {
                //图片尺寸适应单元格
                int adjustBaseSize = 2;
                var adjustImageSize = GetAdjustImageSize(image, cellColumnWidthInPix, cellRowHeightInPix);
                adjustImageWidthInPix = adjustImageSize.Item1 - adjustBaseSize;
                adjustImageHeightInPix = adjustImageSize.Item2 - adjustBaseSize;
            }
            //设置为居中显示
            int columnOffsetPixels = (int)((cellColumnWidthInPix - adjustImageWidthInPix) / 2.0);
            int rowOffsetPixels = (int)((cellRowHeightInPix - adjustImageHeightInPix) / 2.0);
            picture.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            picture.SetSize(adjustImageWidthInPix, adjustImageHeightInPix);
            picture.SetPosition(rowNum - 1, rowOffsetPixels, columnNum - 1, columnOffsetPixels);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="worksheet">worksheet</param>
        /// <param name="imageBytes">图片字节数据</param>
        /// <param name="rowNum">行号,从1开始</param>
        /// <param name="columnNum">列号,从1开始</param>
        /// <param name="autofit">是否自适应单元格</param>
        public static void InsertImage(this ExcelWorksheet worksheet, byte[] imageBytes, int rowNum, int columnNum, bool autofit)
        {
            InsertImage(worksheet, imageBytes, rowNum, columnNum, autofit, new EPPlusConfig());
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="range"></param>
        /// <param name="imgUrl"></param>
        /// <param name="width">字符</param>
        /// <param name="height">磅</param>
        /// <param name="margin"></param>
        public static void SetImage(this ExcelRange range, string imgUrl, double width = 0, double height = 0, double margin = 0)
        {
            range.Worksheet.SetImage(range.Start.Row, range.Start.Column, imgUrl, width, height, margin);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="address"></param>
        /// <param name="imgUrl"></param>
        /// <param name="width">字符</param>
        /// <param name="height">磅</param>
        /// <param name="margin"></param>
        public static void SetImage(this ExcelWorksheet sheet, string address, string imgUrl, double width = 0, double height = 0, double margin = 0)
        {
            var innerAddress = new ExcelAddress(address);
            sheet.SetImage(innerAddress.Start.Row, innerAddress.Start.Column, imgUrl, width, height, margin);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="imgUrl"></param>
        /// <param name="width">字符</param>
        /// <param name="height">磅</param>
        /// <param name="margin"></param>
        public static void SetImage(this ExcelWorksheet sheet, int row, int column, string imgUrl, double width = 0, double height = 0, double margin = 0)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(imgUrl))
                {
                    return;
                }
                imgUrl = imgUrl.Trim();
                using var client = new HttpClient();
                var stream = client.GetStreamAsync(imgUrl).GetAwaiter().GetResult();
                // epplus bug
                // https://github.com/JanKallman/EPPlus/issues/291
                // 解析图片
                using Bitmap img = Image.FromStream(stream) as Bitmap;
                if (img.HorizontalResolution == 0 || img.VerticalResolution == 0)
                {
                    img.SetResolution(96, 96);
                }
                using ExcelPicture pic = sheet.Drawings.AddPicture($"image_{DateTime.Now.Ticks}", img);

                // 当外部传入了有效的 width 和 height 值
                if (width * height > 0)
                {
                    sheet.Row(row).Height = PixelUtil.CellPixelWidth(height + margin * 2);
                    sheet.Column(column).Width = PixelUtil.CellPixelHeight(width + margin * 2);
                }
                else
                {
                    height = sheet.Row(row).Height - margin * 2;
                    width = sheet.Column(column).Width - margin * 2;
                }

                // 设置图片大小
                pic.SetSize((int)width, (int)height);
                // 设置图片放置位置
                pic.SetPosition(row - 1, (int)margin, column - 1, (int)margin);
            }
            catch (Exception)
            {
                //Console.WriteLine($"{DateTime.Now} 插入图片失败：{imgUrl}\r\n{ex.Message} \r\n{ex.StackTrace}");
            }
        }

        #endregion

        #region 读取数据

        /// <summary>
        /// 读取合并单元格的值
        /// </summary>
        /// <param name="worksheet">worksheet</param>
        /// <param name="rowNum">行号,从1开始</param>
        /// <param name="columnNum">列号,从1开始</param>
        /// <returns></returns>
        public static object GetMergeCellValue(this ExcelWorksheet worksheet, int rowNum, int columnNum)
        {
            string range = worksheet.MergedCells[rowNum, columnNum];
            if (range == null)
            {
                return worksheet.Cells[rowNum, columnNum].Value;
            }
            else
            {
                return worksheet.Cells[(new ExcelAddress(range)).Start.Row, (new ExcelAddress(range)).Start.Column].Value;
            }
        }

        /// <summary>
        /// 读取Excel数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">worksheet</param>
        /// <param name="startRow">开始行号</param>
        /// <param name="endRow">结束行号</param>
        /// <returns></returns>
        public static List<T> ReadDatas<T>(this ExcelWorksheet worksheet, int startRow = 2, int? endRow = null) where T : class, new()
        {
            var result = new List<T>();
            var dataType = typeof(T);
            var properties = dataType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var excelColumnProperties = properties.Where(p => p.GetCustomAttributes<ExcelColumnNameAttribute>().Any()).ToArray();
            int maxReadRow = worksheet.Dimension?.End?.Row ?? 0;
            if(endRow != null)
            {
                maxReadRow = Math.Min(endRow.Value, maxReadRow);
            }
            for (int row = startRow; row <= maxReadRow; row++)
            {
                bool hasRowData = false;
                var checkPropertyErrorMessages = new List<string>();
                //创建对象实例
                var dataInstance = Activator.CreateInstance(dataType) as T;
                foreach (var excelColumnProperty in excelColumnProperties)
                {
                    object value = null;
                    //Excel列名称特性
                    var columnNameAttr = excelColumnProperty.GetCustomAttribute<ExcelColumnNameAttribute>();
                    string columnName = columnNameAttr?.ColumnName;
                    if (string.IsNullOrEmpty(columnName))
                    {
                        continue;
                    }
                    //检查excel列名合法性, 合法excel列名,如: B, AB, AZ
                    ExcelAddressUtil.CheckColumnName(columnName);

                    //判断属性是否是可空类型
                    var propertyType = excelColumnProperty.PropertyType;
                    if (IsNullableType(propertyType))
                    {
                        propertyType = Nullable.GetUnderlyingType(propertyType);
                    }

                    try
                    {
                        //excel单元格值验证
                        value = worksheet.Cells[row, ExcelAddressUtil.Col(columnName)].Value;
                        var validationAttributes = excelColumnProperty.GetCustomAttributes()
                            .Where(a => !a.GetType().IsAbstract && a is ValidationAttribute)
                            .Select(a => a as ValidationAttribute);
                        foreach (var validationAttribute in validationAttributes)
                        {
                            if (!validationAttribute.IsValid(value))
                            {
                                string errorMessage = validationAttribute.FormatErrorMessage(excelColumnProperty.Name);
                                checkPropertyErrorMessages.Add($"{errorMessage} at excel row:{row}, column:{columnName};");
                            }
                        }

                        //对象实例赋值
                        var text = worksheet.Cells[row, ExcelAddressUtil.Col(columnName)].Text;
                        if (string.IsNullOrEmpty(text))
                        {
                            continue;
                        }
                        object setValue = text;
                        if (propertyType == typeof(string))
                        {
                            //字符串去掉前后空格
                            setValue = text.Trim();
                        }
                        setValue = Convert.ChangeType(text, propertyType);
                        excelColumnProperty.SetValue(dataInstance, setValue);
                        hasRowData = true;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"assignment failure - {ex.Message}, at excel row:{row}, column:{columnName}, value:{value}", ex);
                    }
                }

                if (hasRowData)
                {
                    if (checkPropertyErrorMessages.Any())
                    {
                        throw new Exception(string.Join(Environment.NewLine, checkPropertyErrorMessages));
                    }
                    result.Add(dataInstance);
                }
            }
            return result;
        }

        #endregion

        #region 图表

        /// <summary>
        /// 设置饼图 datalabel 百分比的数字格式，可用于设置保留小数位数
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="formatString"></param>
        public static void DataLabelPercentageFormat(this ExcelChart chart, string formatString = "0.00%")
        {
            var xdoc = chart.ChartXml;
            var nsuri = xdoc.DocumentElement.NamespaceURI;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);
            nsm.AddNamespace("c", nsuri);

            var numFmtNode = xdoc.CreateElement("c:numFmt", nsuri);
            var formatCodeAtt = xdoc.CreateAttribute("formatCode", nsuri);
            formatCodeAtt.Value = formatString;
            numFmtNode.Attributes.Append(formatCodeAtt);

            var sourceLinkedAtt = xdoc.CreateAttribute("sourceLinked", nsuri);
            sourceLinkedAtt.Value = "0";
            numFmtNode.Attributes.Append(sourceLinkedAtt);
            var dLblsNode = xdoc.SelectSingleNode("c:chartSpace/c:chart/c:plotArea/c:pieChart/c:dLbls", nsm);
            dLblsNode.AppendChild(numFmtNode);
        }
       
        #endregion

        #region private methods

        /// <summary>
        /// 自动填充worksheet数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="title"></param>
        /// <param name="datas"></param>
        private static void AutoFillData<T>(ExcelWorksheet worksheet, string title, IEnumerable<T> datas) where T : class
        {
            var properties = typeof(T).GetProperties();
            properties = properties.Where(p => p.CustomAttributes.Any(x => x.AttributeType == typeof(ExcelColumnNameAttribute))).ToArray();
            if (properties.Length <= 0)
            {
                return;
            }

            int increaseRow = 0;
            //title
            if (!string.IsNullOrEmpty(title))
            {
                int titleRow = 1;
                worksheet.Cells[titleRow, 1, titleRow, properties.Length].Style.Font.Bold = true;
                worksheet.Cells[titleRow, 1, titleRow, properties.Length].Style.Font.Size = 24.0F;
                worksheet.Cells[titleRow, 1, titleRow, properties.Length].Merge = true;
                worksheet.Cells[titleRow, 1].Value = title;
                worksheet.Row(titleRow).Height = 40.0D;
                worksheet.View.FreezePanes(titleRow + 1, 1);
                increaseRow++;
            }

            //header
            int headerRow = 1 + increaseRow;
            worksheet.Cells[headerRow, 1, headerRow, properties.Length].Style.Font.Bold = true;
            worksheet.Row(headerRow).Height = 30.0D;
            worksheet.View.FreezePanes(headerRow + 1, 1);
            for (int i = 0; i < properties.Length; i++)
            {
                string header = properties[i].Name;
                var columnNameProperty = properties[i].GetCustomAttributes(typeof(ExcelColumnNameAttribute), false).FirstOrDefault();
                if (columnNameProperty != null)
                {
                    header = ((ExcelColumnNameAttribute)columnNameProperty).ColumnName;
                    header = !string.IsNullOrEmpty(header) ? header : properties[i].Name;
                }
                worksheet.Cells[headerRow, i + 1].Value = header;
                worksheet.Column(i + 1).Width = (header.Length + 2) * 2;
            }

            //datas
            int startRow = 2 + increaseRow;
            for(int j = 0; j < datas.Count(); j++)
            {
                var data = datas.ElementAt(j);
                for (int k = 0; k < properties.Length; k++)
                {
                    var property = properties[k];
                    object value = property.GetValue(data, null);
                    WriteToCelll(worksheet, startRow + j, k + 1, value, new EPPlusConfig() { AutoFitColumn = true });
                    worksheet.Row(startRow + j).Height = 30.0D;
                }
            }

            worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.WrapText = true;
        }

        /// <summary>
        /// 将数据写入单元格中
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <param name="config"></param>
        private static void WriteToCelll(ExcelWorksheet worksheet, int row, int column, object value, EPPlusConfig config)
        {
            var excelRange = worksheet.Cells[row, column];
            if (value == null)
            {
                excelRange.Value = string.Empty;
                return;
            }

            var type = value.GetType();
            //值为日期类型
            if (type == typeof(DateTime))
            {
                excelRange.Style.Numberformat.Format = config.DateTimeFormat;
            }
            //值为整数类型
            else if (type == typeof(int) || type == typeof(long))
            {
                excelRange.Style.Numberformat.Format = "#0";
            }
            //值为字节数组类型,则默认为图片数据进行处理
            else if (type == typeof(byte[]))
            {
                worksheet.Column(column).Width = config.DefaultSizeWithPicture.Item1;
                worksheet.Row(row).Height = config.DefaultSizeWithPicture.Item2;
                InsertImage(worksheet, (byte[])value, row, column, true, new EPPlusConfig() { WidthPixelsBase = 8.0 });
                value = null;
            }
            excelRange.Value = value ?? string.Empty;

            //列宽自适应
            double setColumnWidth = worksheet.Column(column).Width;
            if (config.AutoFitColumn && !string.IsNullOrEmpty(value?.ToString()))
            {
                double originColumnWidth = worksheet.Column(column).Width;
                double newColumnWidth = (value.ToString().Length + 2) * 2;
                if (newColumnWidth > originColumnWidth)
                {
                    setColumnWidth = newColumnWidth;
                }
                //worksheet.Column(column).AutoFit();
            }
            setColumnWidth = setColumnWidth > config.MaxColumnWidth ? config.MaxColumnWidth : setColumnWidth;
            worksheet.Column(column).Width = setColumnWidth;

            return;
        }

        /// <summary>
        /// 获取模板字段字典(key:所在列, value:字段名称)
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        private static Dictionary<int, string> GetTemplateFieldDic(ExcelWorksheet worksheet, EPPlusConfig config)
        {
            var result = new Dictionary<int, string>();
            var start = worksheet.Dimension?.Start;
            var end = worksheet.Dimension?.End;
            if (start == null || end == null)
            {
                return result;
            }

            var propertyRegex = new Regex(config.PropertyMatchRegex);
            var excelRange = worksheet.Cells[start.Row, start.Column, end.Row, end.Column];
            var excelCells = excelRange.Where(e => e.Value != null && !string.IsNullOrWhiteSpace(e.Value.ToString()) && propertyRegex.IsMatch(e.Value.ToString())).ToList();
            if (excelCells == null || excelCells.Count <= 0)
            {
                return result;
            }

            foreach (var cell in excelCells)
            {
                var value = cell.Value;
                var cellValue = cell.Value.ToString().Trim();
                var match = propertyRegex.Match(cellValue);
                if(match.Success && match.Groups.Count >= 2)
                {
                    string propertyName = match.Groups[1].Value.Trim();
                    if (!result.ContainsKey(cell.Start.Column))
                    {
                        result.Add(cell.Start.Column, propertyName);
                    }
                }
                cell.Value = string.Empty;
            }
            return result;
        }

        /// <summary>
        /// 获取自适应调整后的图片尺寸
        /// </summary>
        /// <param name="image"></param>
        /// <param name="cellColumnWidthInPix"></param>
        /// <param name="cellRowHeightInPix"></param>
        /// <returns>item1:调整后的图片宽度; item2:调整后的图片高度</returns>
        private static Tuple<int, int> GetAdjustImageSize(Image image, int cellColumnWidthInPix, int cellRowHeightInPix)
        {
            int imageWidthInPix = image.Width;
            int imageHeightInPix = image.Height;
            //调整图片尺寸,适应单元格
            int adjustImageWidthInPix;
            int adjustImageHeightInPix;
            if (imageHeightInPix * cellColumnWidthInPix > imageWidthInPix * cellRowHeightInPix)
            {
                //图片高度固定,宽度自适应
                adjustImageHeightInPix = cellRowHeightInPix;
                double ratio = (1.0) * adjustImageHeightInPix / imageHeightInPix;
                adjustImageWidthInPix = (int)(imageWidthInPix * ratio);
            }
            else
            {
                //图片宽度固定,高度自适应
                adjustImageWidthInPix = cellColumnWidthInPix;
                double ratio = (1.0) * adjustImageWidthInPix / imageWidthInPix;
                adjustImageHeightInPix = (int)(imageHeightInPix * ratio);
            }
            return new Tuple<int, int>(adjustImageWidthInPix, adjustImageHeightInPix);
        }

        /// <summary>
        /// 判断是否为可空类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static bool IsNullableType(Type type)
        {
            return (type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>)));
        }

        #endregion 
    }
}
