using Edesoft.APP.Abstractions.Excel;
using Edesoft.APP.Abstractions.IO;
using Edesoft.APP.Abstractions.Pdf;
using Edesoft.APP.Infrastructure.Configuration;

namespace Edesoft.APP.Worker
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IIOApplication _ioApplication;
        private readonly IPdfApplication _parserApplication;
        private readonly IExcelApplication _excelApplication;

        public Worker(
            ILogger<Worker> logger,
            IIOApplication ioApplication,
            IPdfApplication parserApplication,
            IExcelApplication excelApplication)
        {
            _logger = logger;
            _ioApplication = ioApplication;
            _parserApplication = parserApplication;
            _excelApplication = excelApplication;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            await Task.Delay(1000, stoppingToken);

            _logger.LogInformation(string.Format("{0} - Iniciando serviço.",
                DateTimeOffset.Now));

            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation(string.Format("{0} - Processo iniciado.",
                    DateTimeOffset.Now));

                try
                {
                    var files = _ioApplication.GetFiles();

                    if (files.Length == 0)
                    {
                        _logger.LogInformation(string.Format("{0} - Não há arquivos para serem processados.",
                            DateTimeOffset.Now));
                    }
                    else
                    {
                        _logger.LogInformation(string.Format("{0} - {1} PDF{2} para leitura.",
                            DateTimeOffset.Now, files.Length, files.Length > 1 ? "s" : ""));

                        List<Dictionary<string, string>> listValues = new();

                        foreach (var file in files)
                        {
                            try
                            {
                                var pdfToText = _parserApplication.ConvertPdfToText(file);

                                if (pdfToText.Length > 0)
                                {
                                    var values = _excelApplication.GenerateExcelFromText(pdfToText, file);

                                    if (values.Keys.Count > 0)
                                        listValues.Add(values);
                                }
                            }
                            catch (Exception e)
                            {
                                _logger.LogWarning(string.Format("{0} - {1}.",
                                    DateTimeOffset.Now, e.Message));
                            }
                        }

                        _excelApplication.GenerateExcel(listValues, Settings.Configuration["IO:Processed"] ??
                            throw new ArgumentNullException("IO:Processed"));
                    }
                }
                catch (Exception e)
                {
                    _logger.LogError(string.Format("{0} - {1}.",
                        DateTimeOffset.Now, e.Message));
                }
                finally
                {
                    _logger.LogInformation(string.Format("{0} - Aguardando próxima execução.",
                            DateTimeOffset.Now));

                    await Task.Delay(int.Parse(Settings.Configuration["WorkerConfig:Delay"] ??
                        throw new ArgumentNullException("Delay")), stoppingToken);
                }
            }
        }
    }
}