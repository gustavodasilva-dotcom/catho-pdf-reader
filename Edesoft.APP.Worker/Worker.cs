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
                    var toProcessPath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        Settings.Configuration["IO:ToProcess"] ??
                            throw new ArgumentNullException("IO:ToProcess"));

                    var files = _ioApplication.GetFiles(toProcessPath);

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
                                else
                                {
                                    _logger.LogInformation(string.Format("{0} - O PDF está vazio.",
                                        DateTimeOffset.Now));

                                    var errorFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                        Settings.Configuration["IO:Error"] ??
                                            throw new ArgumentNullException("IO:Error"),
                                        DateTime.Now.Year.ToString(),
                                        DateTime.Now.Month.ToString(),
                                        DateTime.Now.Day.ToString());

                                    _ioApplication.MoveTo(errorFolder, file);
                                }
                            }
                            catch (Exception e)
                            {
                                _logger.LogWarning(string.Format("{0} - {1}.",
                                    DateTimeOffset.Now, e.Message));
                            }
                        }

                        var processedFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                            Settings.Configuration["IO:Processed"] ??
                                throw new ArgumentNullException("IO:Processed"));

                        _excelApplication.GenerateExcel(listValues, processedFolder);
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