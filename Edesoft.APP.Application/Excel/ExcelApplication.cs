using Edesoft.APP.Abstractions.Excel;
using Edesoft.APP.Abstractions.IO;
using Edesoft.APP.Infrastructure.Configuration;
using Edesoft.APP.Tools.Extensions;
using IronXL;
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
            local = local.Replace("-", ",");
            var localSplit = local.Split(",");
            return localSplit[0].Trim() + ", " + localSplit[1].Trim();
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
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = workbook.CreateWorkSheet("Planilha1");

            _logger.LogInformation(string.Format("{0} - Criando arquivo Excel.",
                DateTimeOffset.Now));

            var sheetLine = 1;

            sheet[$"A${sheetLine}"].Value = excelHeader[0];
            sheet[$"B${sheetLine}"].Value = excelHeader[1];
            sheet[$"C${sheetLine}"].Value = excelHeader[2];
            sheet[$"D${sheetLine}"].Value = excelHeader[3];
            sheet[$"E${sheetLine}"].Value = excelHeader[4];
            sheet[$"F${sheetLine}"].Value = excelHeader[5];
            sheet[$"G${sheetLine}"].Value = excelHeader[6];

            foreach (var item in valuesToExcel)
            {
                try
                {
                    sheetLine++;

                    sheet[$"A${sheetLine}"].Value = item[excelHeader[0]];
                    sheet[$"B${sheetLine}"].Value = item[excelHeader[1]];
                    sheet[$"C${sheetLine}"].Value = item[excelHeader[2]];
                    sheet[$"D${sheetLine}"].Value = item[excelHeader[3]];
                    sheet[$"E${sheetLine}"].Value = item[excelHeader[4]];
                    sheet[$"F${sheetLine}"].Value = item[excelHeader[5]];
                    sheet[$"G${sheetLine}"].Value = item[excelHeader[6]];

                    _ioApplication.MoveTo(
                        Path.Combine(
                            toPath,
                            DateTime.Now.Year.ToString(),
                            DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString()),
                        item.GetValueOrDefault("file") ??
                            throw new ArgumentNullException("file"));
                }
                catch (Exception e)
                {
                    _logger.LogWarning(string.Format("{0} - {1}.",
                        DateTimeOffset.Now, e.Message));
                }
            }

            var savePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                Settings.Configuration["ExcelInfo:SavePath"] ??
                    throw new ArgumentNullException("ExcelInfo:SavePath"),
                DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(),
                DateTime.Now.Day.ToString());

            _ioApplication.CreateDir(savePath);

            var fileName = $"Curriculos_{DateTime.Now:ddMMyyyyhhmmss}.xlsx";

            savePath = Path.Combine(savePath, fileName);

            _logger.LogInformation(string.Format("{0} - Gerando arquivo {1}.",
                DateTimeOffset.Now, fileName));

            workbook.SaveAs(savePath);

            _logger.LogInformation(string.Format("{0} - Arquivo {1} salvo e gerado.",
                DateTimeOffset.Now, fileName));
        }
    }
}
