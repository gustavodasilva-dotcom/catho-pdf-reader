using Edesoft.APP.Abstractions.Pdf;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Extensions.Logging;

namespace Edesoft.APP.Application.Pdf
{
    public class PdfApplication : IPdfApplication
    {
        private readonly ILogger<PdfApplication> _logger;

        public PdfApplication(ILogger<PdfApplication> logger) =>
            _logger = logger;

        public string ConvertPdfToText(string path)
        {
            PdfReader reader = new(path);
            string pdfToText = string.Empty;

            _logger.LogInformation(string.Format("{0} - Lendo conteúdo do PDF {1}.",
                DateTimeOffset.Now, path.Split("\\").LastOrDefault()));

            for (int page = 1; page <= reader.NumberOfPages; page++)
                pdfToText += PdfTextExtractor.GetTextFromPage(reader, page);

            reader.Close();

            return pdfToText;
        }
    }
}
