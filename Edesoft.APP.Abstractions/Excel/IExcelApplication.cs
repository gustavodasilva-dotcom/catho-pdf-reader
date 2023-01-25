namespace Edesoft.APP.Abstractions.Excel
{
    public interface IExcelApplication
    {
        Dictionary<string, string> GenerateExcelFromText(string pdfToText, string file);
        void GenerateExcel(List<Dictionary<string, string>> valuesToExcel, string toPath);
    }
}
