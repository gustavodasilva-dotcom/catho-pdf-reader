namespace Edesoft.APP.Abstractions.IO
{
    public interface IIOApplication
    {
        string[] GetFiles();
        void MoveTo(string toPath, string sourceFilePath);
        void CreateDir(string path);
    }
}
