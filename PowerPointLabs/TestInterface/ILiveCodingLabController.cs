namespace TestInterface
{
    public interface ILiveCodingLabController
    {
        void OpenPane();

        void InsertTextCodeBox(string text);

        void InsertFileCodeBox(string filePath);
        void InsertDiffCodeBox(string diffFilePath);
    }
}
