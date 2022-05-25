### 封装了EPPlus，提供OfficeOpenXml.ExcelWorksheet类的扩展方法:
- 文件转换:  `ToFile(this ExcelWorksheet worksheet, ExcelToFileType toFileType, string filePath)`

### ExcelToFileType支持的类型如下:
- `ExcelToFileType.Text`: Excel转普通文本文件
- `ExcelToFileType.Csv`: Excel转Csv文件
- `ExcelToFileType.Html`: Excel转Html文件