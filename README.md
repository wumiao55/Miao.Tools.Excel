### Miao.Tools.Excel 提供OfficeOpenXml.ExcelWorksheet类的扩展方法:
- 写入数据:  `WriteData<T>(this ExcelWorksheet worksheet, T data)`
- 写入数据集合:  `WriteDatas<T>(this ExcelWorksheet worksheet, ICollection<T> datas, int startRowNum)`
- 自动填充数据:  `AutoFill<T>(this ExcelWorksheet worksheet, List<T> datas)`
- 插入图片:  `InsertImage(this ExcelWorksheet worksheet, byte[] imageBytes, int rowNum, int columnNum, bool autofit)`
- 读取数据:  `List<T> ReadDatas<T>(this ExcelWorksheet worksheet, int startRow, int endRow)`

### Miao.Tools.Excel.Convertor 提供OfficeOpenXml.ExcelWorksheet类的扩展方法:
- 文件转换:  `ToFile(this ExcelWorksheet worksheet, ExcelToFileType toFileType, string filePath)`
### ExcelToFileType支持的类型如下:
- `ExcelToFileType.Text`: Excel转普通文本文件
- `ExcelToFileType.Csv`: Excel转Csv文件
- `ExcelToFileType.Html`: Excel转Html文件
