using Edesoft.APP.Abstractions.IO;
using Microsoft.Extensions.Logging;

namespace Edesoft.APP.Application.IO
{
    public class IOApplication : IIOApplication
    {
        private readonly ILogger<IOApplication> _logger;

        public IOApplication(ILogger<IOApplication> logger) =>
            _logger = logger;

        public string[] GetFiles(
            string path,
            bool createPath = true)
        {
            if (createPath)
                CreateDir(path);

            return Directory.GetFiles(path);
        }

        public void MoveTo(
            string toPath,
            string sourceFilePath,
            bool renameFile = true)
        {
            var fileName = sourceFilePath.Split('\\').Last();

            if (renameFile)
                fileName = RenameFile(fileName);

            CreateDir(toPath);

            toPath = Path.Combine(toPath, fileName);

            File.Move(@sourceFilePath, @toPath);
        }

        public string RenameFile(string fileName)
        {
            var fileNameSplited = fileName.Split(".");

            fileName = string.Concat(fileNameSplited[0],
                "_", DateTime.Now.ToString("ddMMyyyyHHssff"), ".",
                fileNameSplited[fileNameSplited.Length - 1]);

            return fileName;
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
