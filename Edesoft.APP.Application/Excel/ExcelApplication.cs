using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Edesoft.APP.Abstractions.Excel;
using Edesoft.APP.Abstractions.IO;
using Edesoft.APP.Infrastructure.Configuration;
using Edesoft.APP.Tools.Extensions;
using Microsoft.Extensions.Logging;

namespace Edesoft.APP.Application.Excel
{
    public class ExcelApplication : IExcelApplication
    {
        private readonly string[] excelHeader = {
            "nome",
            "local",
            "celular",
            "email",
            "linkedIn",
            "salario",
            "comosoubedavaga",
            "file"
        };

        private readonly IIOApplication _ioApplication;
        private readonly ILogger<ExcelApplication> _logger;

        public ExcelApplication(
            IIOApplication ioApplication,
            ILogger<ExcelApplication> logger)
        {
            _ioApplication = ioApplication;
            _logger = logger;
        }

        private static string GetName(string textLine, string file)
        {
            try
            {
                return textLine.Substring(0, textLine.IndexOf("(") - 1).Trim();
            }
            catch (Exception)
            {
                var nameFromFile = file.Split("\\").Last();
                return nameFromFile.Split(".").First();
            }
        }

        public static string GetEmail(string textLine)
        {
            return textLine.Substring(textLine.IndexOf(":") + 1).Trim();
        }

        private static string GetCellphoneNumber(string textLine)
        {
            return textLine.Substring(textLine.IndexOf(":") + 1).Trim();
        }

        private static string GetSalary(string textLine)
        {
            return textLine.Substring(textLine.IndexOf("$") + 1).Trim();
        }

        private static string GetCity(string textLine)
        {
            var local = textLine.Substring(textLine.IndexOf(":") + 1).Trim();
            return local.Split("-")[0];
        }

        public Dictionary<string, string> GenerateExcelFromText(
            string pdfToText,
            string file)
        {
            _logger.LogInformation(string.Format("{0} - Extraíndo dados do arquivo {1}.",
                DateTimeOffset.Now, file.Split("\\").Last()));

            Dictionary<string, string> valuesToExcel = new();

            var linkedin = Settings.Configuration["ExcelInfo:Linkedin"] ??
                throw new ArgumentNullException("ExcelInfo:Linkedin");
            valuesToExcel.Add(excelHeader[4], linkedin);

            var comosoubedavaga = Settings.Configuration["ExcelInfo:ComoSoubeVaga"] ??
                throw new ArgumentNullException("ExcelInfo:ComoSoubeVaga");
            valuesToExcel.Add(excelHeader[6], comosoubedavaga);

            valuesToExcel.Add(nameof(file), file);

            foreach (var line in pdfToText.Split("\n"))
            {
                if (line.Contains("Data da última alteração"))
                {
                    var nome = GetName(line, file);

                    if (!string.IsNullOrEmpty(nome))
                        valuesToExcel.Add(excelHeader[0], nome);
                }
                else if (line.Contains("E-mail"))
                {
                    var email = GetEmail(line);

                    if (!string.IsNullOrEmpty(email))
                        valuesToExcel.Add(excelHeader[3], email);
                }
                else if (line.Contains("Cidade"))
                {
                    var local = GetCity(line);

                    if (!string.IsNullOrEmpty(local))
                        valuesToExcel.Add(excelHeader[1], local);
                }
                else if (line.Contains("Telefone celular"))
                {
                    var celular = GetCellphoneNumber(line);

                    if (!string.IsNullOrEmpty(celular))
                        valuesToExcel.Add(excelHeader[2], celular);
                }
                else if (line.Contains("A Combinar"))
                {
                    valuesToExcel.Add(excelHeader[5], "0");
                }
                else if (line.Contains("R$"))
                {
                    var salario = GetSalary(line);

                    if (!string.IsNullOrEmpty(salario))
                        valuesToExcel.Add(excelHeader[5], salario.FormatRealToExcelForm());
                }

                if (valuesToExcel.Count == excelHeader.Length)
                    break;
            }

            _logger.LogInformation(string.Format("{0} - Extração concluída.",
                DateTimeOffset.Now));

            return valuesToExcel;
        }

        public void GenerateExcel(
            List<Dictionary<string, string>> valuesToExcel,
            string toPath)
        {
            _logger.LogInformation(string.Format("{0} - Criando arquivo Excel.",
                DateTimeOffset.Now));

            var savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                Settings.Configuration["IO:ExcelSavePath"] ??
                    throw new ArgumentNullException("IO:ExcelSavePath"));

            var files = _ioApplication.GetFiles(savePath);

            if (files.Length > 0)
            {
                foreach (var file in files)
                {
                    var oldFilePath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        Settings.Configuration["IO:OldFile"] ??
                            throw new ArgumentNullException("IO:OldFile"));

                    _ioApplication.MoveTo(oldFilePath, file);
                }
            }

            savePath = Path.Combine(savePath,
                Settings.Configuration["ExcelInfo:FileName"] ??
                    throw new ArgumentNullException("ExcelInfo:FileName"));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(
                savePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Planilha1"
                };

                sheets.Append(sheet);

                Row headerRow = new();

                foreach (var header in excelHeader)
                {
                    if (header == excelHeader[excelHeader.Length - 1])
                        continue;

                    var cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(header)
                    };

                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (var item in valuesToExcel)
                {
                    try
                    {
                        Row newRow = new();

                        foreach (var header in excelHeader)
                        {
                            try
                            {
                                if (header == excelHeader[excelHeader.Length - 1])
                                    continue;

                                var cell = new Cell
                                {
                                    DataType = CellValues.String,
                                    CellValue = new CellValue(item[header])
                                };

                                newRow.AppendChild(cell);
                            }
                            catch (Exception e)
                            {
                                _logger.LogWarning(string.Format("{0} - {1}.",
                                    DateTimeOffset.Now, e.Message));

                                throw;
                            }
                        }

                        sheetData.AppendChild(newRow);

                        var processedFolder = Path.Combine(toPath,
                            DateTime.Now.Year.ToString(),
                            DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString());

                        _ioApplication.MoveTo(processedFolder,
                            item.GetValueOrDefault("file") ??
                                throw new ArgumentNullException("file"));
                    }
                    catch (Exception)
                    {
                        var errorFolder = Path.Combine(
                            AppDomain.CurrentDomain.BaseDirectory,
                            Settings.Configuration["IO:Error"] ??
                                throw new ArgumentNullException("IO:Error"),
                            DateTime.Now.Year.ToString(),
                            DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString());

                        _ioApplication.MoveTo(errorFolder,
                            item.GetValueOrDefault("file") ??
                                throw new ArgumentNullException("file"));
                    }
                }

                _logger.LogInformation(string.Format("{0} - Gerando arquivo {1}.",
                    DateTimeOffset.Now, savePath));

                workbookPart.Workbook.Save();

                _logger.LogInformation(string.Format("{0} - Arquivo {1} salvo e gerado.",
                    DateTimeOffset.Now, savePath));
            }
        }
    }
}
