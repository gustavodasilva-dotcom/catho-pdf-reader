using Edesoft.APP.Abstractions.IO;
using Edesoft.APP.Infrastructure.Configuration;
using Microsoft.Extensions.Logging;

namespace Edesoft.APP.Application.IO
{
    public class IOApplication : IIOApplication
    {
        private readonly ILogger<IOApplication> _logger;

        public IOApplication(ILogger<IOApplication> logger) =>
            _logger = logger;

        public string[] GetFiles()
        {
            var toProcessDir = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                Settings.Configuration["IO:ToProcess"] ??
                    throw new ArgumentNullException("IO:ToProcess"),
                DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(),
                DateTime.Now.Day.ToString());

            CreateDir(toProcessDir);

            return Directory.GetFiles(toProcessDir);
        }

        public void MoveTo(string toPath, string sourceFilePath)
        {
            var fileName = sourceFilePath.Split('\\').Last();

            toPath = AppDomain.CurrentDomain.BaseDirectory + toPath;
            var toPathLogger = toPath;

            CreateDir(toPath);

            toPath = Path.Combine(toPath, fileName);

            _logger.LogInformation(
                string.Format("{0} - Movendo arquivo {1} de {2} para {3}.",
                DateTimeOffset.Now,
                fileName,
                sourceFilePath.Split("\\")[sourceFilePath.Split("\\").Length - 5],
                toPathLogger.Split("\\")[toPathLogger.Split("\\").Length - 4]));

            File.Move(@sourceFilePath, @toPath);

            _logger.LogInformation(string.Format("{0} - Arquivo {1} movido.",
                DateTimeOffset.Now, fileName));
        }

        public void CreateDir(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);

                _logger.LogInformation(string.Format("{0} - Criando pastas {1}.",
                    DateTimeOffset.Now, path));
            }
        }
    }
}
