using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Text;

namespace excel_to_dat_file.Controllers
{
    public class UploadController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [RequestSizeLimit(100_000_000)] // 100 MB
        public async Task<IActionResult> ExcelToDat(IFormFile file, bool hasHeader = true)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            if (Path.GetExtension(file.FileName).ToLowerInvariant() != ".xlsx")
                return BadRequest("Only .xlsx files are supported.");

            int expectedColumnCount = 49;

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Position = 0;

            using var workbook = new XLWorkbook(ms);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
                return BadRequest("Workbook contains no worksheets.");

            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
                return BadRequest("Worksheet is empty.");

            var firstRow = usedRange.FirstRowUsed().RowNumber();
            var lastRow = usedRange.LastRowUsed().RowNumber();
            var lastColumn = usedRange.LastColumnUsed().ColumnNumber();

            if (lastColumn < expectedColumnCount)
                return BadRequest($"Invalid template. Expected {expectedColumnCount} columns.");

            var startRow = hasHeader ? firstRow + 3 : firstRow;
            if (startRow > lastRow)
                return BadRequest("No data rows found.");

            var sb = new StringBuilder();

            // Header row
            var headerRow = worksheet.Row(startRow);

            string form = headerRow.Cell(2).GetValue<string>() ?? string.Empty;
            string tinNo = headerRow.Cell(3).GetValue<string>() ?? string.Empty;
            string tinNoExtension = headerRow.Cell(4).GetValue<string>() ?? string.Empty;
            string dateValue = headerRow.Cell(5).DataType == XLDataType.DateTime
                ? headerRow.Cell(5).GetDateTime().ToString("MM/dd/yyyy")
                : headerRow.Cell(5).GetValue<string>() ?? string.Empty;

            string headerLine = $"H{form},{tinNo},{tinNoExtension},{dateValue}";
            sb.AppendLine(headerLine);

            for (int r = startRow; r <= lastRow; r++)
            {
                var row = worksheet.Row(r);
                if (r == lastRow) expectedColumnCount = 36;

                var values = new string[expectedColumnCount];

                for (int c = 1; c <= expectedColumnCount; c++)
                {
                    var cell = row.Cell(c);
                    int columnNumber = cell.Address.ColumnNumber;

                    bool isQuotedColumn = columnNumber > 8 && columnNumber < 12 && r != lastRow;

                    string value = cell.DataType switch
                    {
                        XLDataType.DateTime => cell.GetDateTime().ToString("MM/dd/yyyy"),
                        XLDataType.Number => cell.GetDouble().ToString("0.00"),
                        _ => cell.GetValue<string>() ?? string.Empty
                    };

                    value = value
                        .Replace("\r", " ")
                        .Replace("\n", " ")
                        .Trim();

                    values[c - 1] = isQuotedColumn
                        ? $"\"{value}\""
                        : value;
                }


                sb.AppendLine(string.Join(",", values));
            }

            var datBytes = Encoding.UTF8.GetBytes(sb.ToString());
            var cleanDate = new string(dateValue.Where(char.IsDigit).ToArray());

            var outputFileName = $"{tinNo}{tinNoExtension}{cleanDate}{form}.dat";


            return File(datBytes, "application/octet-stream", outputFileName);
        }
    
        [HttpPost]
        [RequestSizeLimit(100_000_000)] // 100 MB
        public async Task<IActionResult> ExcelToDatV2(IFormFile file, bool hasHeader = true)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            if (Path.GetExtension(file.FileName).ToLowerInvariant() != ".xlsx")
                return BadRequest("Only .xlsx files are supported.");

            int expectedColumnCount = 44;

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Position = 0;

            using var workbook = new XLWorkbook(ms);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
                return BadRequest("Workbook contains no worksheets.");

            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
                return BadRequest("Worksheet is empty.");

            var firstRow = usedRange.FirstRowUsed().RowNumber();
            var lastRow = usedRange.LastRowUsed().RowNumber();
            var lastColumn = usedRange.LastColumnUsed().ColumnNumber();

            if (lastColumn < expectedColumnCount)
                return BadRequest($"Invalid template. Expected {expectedColumnCount} columns.");

            var startRow = hasHeader ? firstRow + 3 : firstRow;
            if (startRow > lastRow)
                return BadRequest("No data rows found.");

            var sb = new StringBuilder();

            // Read Header (Row 2)
            var headerRow = worksheet.Row(2);

            string recordType = headerRow.Cell(1).GetValue<string>().Trim();
            string form = headerRow.Cell(2).GetValue<string>().Trim();
            string tinNo = headerRow.Cell(3).GetValue<string>().Trim();
            string tinExt = headerRow.Cell(4).GetValue<string>().Trim();

            string dateValue = headerRow.Cell(5).DataType == XLDataType.DateTime
                ? headerRow.Cell(5).GetDateTime().ToString("MM/dd/yyyy")
                : headerRow.Cell(5).GetValue<string>().Trim();

            string headerLine = $"H{form},{tinNo},{tinExt},{dateValue}";
            sb.AppendLine(headerLine);

            // This prefix will be added to EVERY D1 row
            string d1Prefix = string.Join(",", recordType, form, tinNo, tinExt, dateValue);

            // Prepare totals (adjust number of totals depending on your columns)
            var totals = new List<decimal>();

            // prepare totals only for amount columns
            var amountTotals = new Dictionary<int, decimal>();

            for (int r = startRow; r <= lastRow; r++)
            {
                var row = worksheet.Row(r);

                if (row.CellsUsed().Count() == 0)
                    continue;

                var values = new List<string>();
                values.Add(d1Prefix);

                for (int c = 1; c <= row.LastCellUsed().Address.ColumnNumber; c++)
                {
                    var cell = row.Cell(c);
                    int columnNumber = cell.Address.ColumnNumber;

                    bool isQuotedColumn = columnNumber >= 4 && columnNumber <= 6;

                    string value = cell.DataType switch
                    {
                        XLDataType.DateTime => cell.GetDateTime().ToString("MM/dd/yyyy"),
                        XLDataType.Number => cell.GetDouble().ToString("0.00"),
                        _ => cell.GetValue<string>() ?? string.Empty
                    };

                    value = value.Replace("\r", " ").Replace("\n", " ").Trim();
                    values.Add(isQuotedColumn ? $"\"{value}\"" : value);

                    // accumulate only amount columns 
                    if (cell.DataType == XLDataType.Number && columnNumber >= 8) // adjust start col
                    {
                        if (!amountTotals.ContainsKey(columnNumber))
                            amountTotals[columnNumber] = 0m;

                        amountTotals[columnNumber] += (decimal)cell.GetDouble();
                    }
                }

                sb.AppendLine(string.Join(",", values));
            }

            //Created C1 LINE 
            var c1Values = new List<string>
                {
                    "C1",
                    form,
                    tinNo,
                    tinExt,
                    dateValue
                };

            // Append totals in column order
            foreach (var kvp in amountTotals.OrderBy(k => k.Key))
            {
                c1Values.Add(kvp.Value.ToString("0.00"));
            }

            sb.AppendLine(string.Join(",", c1Values));


            var datBytes = Encoding.UTF8.GetBytes(sb.ToString());
            var cleanDate = new string(dateValue.Where(char.IsDigit).ToArray());

            var outputFileName = $"{tinNo}{tinExt}{cleanDate}{form}.dat";


            return File(datBytes, "application/octet-stream", outputFileName);
        }
    }
}
