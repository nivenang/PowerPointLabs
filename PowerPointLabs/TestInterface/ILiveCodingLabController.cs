namespace TestInterface
{
    public interface ILiveCodingLabController
    {
        void OpenPane();

        void InsertTextCodeBox(string text, string shapeName);

        void InsertFileCodeBox(string filePath);

        void RefreshTextCodeBox(string oldText, string newText);

        void InsertDiffCodeBox(string diffFilePath);

        void AnimateLineDiff();

        void AnimateBlockDiff();

        void AnimateCharDiff();
    }
}
