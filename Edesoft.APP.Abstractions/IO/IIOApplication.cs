namespace Edesoft.APP.Abstractions.IO
{
    public interface IIOApplication
    {
        string[] GetFiles(
            string path,
            bool createPath = true);

        void MoveTo(
            string toPath,
            string sourceFilePath,
            bool renameFile = true);

        string RenameFile(string fileName);

        void CreateDir(string path);
    }
}
