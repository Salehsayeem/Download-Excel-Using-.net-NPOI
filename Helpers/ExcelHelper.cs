using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DownloadExcel.Helpers
{
    public class ExcelHelper
    {
        public static byte[] CreateFile<T>(List<T> source, string fileType)
        {
            IWorkbook workbook;

            if (fileType == "xlsx")
            {
                workbook = new XSSFWorkbook();
                CreateXlsxContent(workbook, source);
            }
            else if (fileType == "csv")
            {
                workbook = new XSSFWorkbook();
                return CreateCsvContent(source);
            }
            else
            {
                throw new NotSupportedException($"File type '{fileType}' is not supported.");
            }

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }

        private static void CreateXlsxContent<T>(IWorkbook workbook, List<T> source)
        {
            var sheet = workbook.CreateSheet("Sheet1");
            var rowHeader = sheet.CreateRow(0);

            var properties = typeof(T).GetProperties();

            // header
            var font = workbook.CreateFont();
            font.IsBold = true;
            var style = workbook.CreateCellStyle();
            style.SetFont(font);

            var colIndex = 0;
            foreach (var property in properties)
            {
                var cell = rowHeader.CreateCell(colIndex);
                cell.SetCellValue(property.Name);
                cell.CellStyle = style;
                colIndex++;
            }
            // end header

            // content
            var rowNum = 1;
            foreach (var item in source)
            {
                var rowContent = sheet.CreateRow(rowNum);

                var colContentIndex = 0;
                foreach (var property in properties)
                {
                    var cellContent = rowContent.CreateCell(colContentIndex);
                    var value = property.GetValue(item, null);

                    if (value == null)
                    {
                        cellContent.SetCellValue("");
                    }
                    else if (property.PropertyType == typeof(string))
                    {
                        cellContent.SetCellValue(value.ToString());
                    }
                    else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                    {
                        cellContent.SetCellValue(Convert.ToInt32(value));
                    }
                    else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                    {
                        cellContent.SetCellValue(Convert.ToDouble(value));
                    }
                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                    {
                        var dateValue = (DateTime)value;
                        cellContent.SetCellValue(dateValue.ToString("yyyy-MM-dd"));
                    }
                    else cellContent.SetCellValue(value.ToString());

                    colContentIndex++;
                }

                rowNum++;
            }
            // end content
        }

        private static byte[] CreateCsvContent<T>(List<T> source)
        {
            using (var memoryStream = new MemoryStream())
            using (var writer = new StreamWriter(memoryStream))
            {
                var properties = typeof(T).GetProperties();
                var header = string.Join(",", properties.Select(p => p.Name));
                writer.WriteLine(header);

                foreach (var item in source)
                {
                    var line = string.Join(",", properties.Select(p =>
                    {
                        var value = p.GetValue(item, null);
                        return value == null
                            ? ""
                            : (p.PropertyType == typeof(DateTime)
                                ? ((DateTime)value).ToString("yyyy-MM-dd")
                                : value.ToString());
                    }));
                    writer.WriteLine(line);
                }

                writer.Flush();
                return memoryStream.ToArray();
            }
        }
    }
}
