using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.LiveCodingLab.Service
{
    public class CodeBoxFileService
    {
        public static string GetCodeFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show(LiveCodingLabText.ErrorInvalidFileName,
                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                return "";
            }

            return File.ReadAllText(filePath);

        }

        public static List<FileDiff> ParseDiff(string diffPath)
        {
            string diffInput = GetCodeFromFile(diffPath);

            if (diffInput == "")
            {
                return null;
            }

            List<FileDiff> diffList = Diff.Parse(diffInput, Environment.NewLine).ToList();
            return diffList;
        }

        public static List<string> ConvertFileDiffToString(FileDiff diffFile)
        {
            List<ChunkDiff> diffChunks = diffFile.Chunks.ToList();
            List<string> diffList = new List<string>();
            string codeTextBefore = "";
            string codeTextAfter = "";

            foreach (ChunkDiff chunk in diffChunks)
            {
                List<LineDiff> diffLines = chunk.Changes.ToList();
                foreach (LineDiff line in diffLines)
                {
                    if (line.Add)
                    {
                        codeTextAfter += AppendLineEnd(line.Content.Trim().Substring(1));
                    }
                    else if (line.Delete)
                    {
                        codeTextBefore += AppendLineEnd(line.Content.Trim().Substring(1));
                    }
                    else
                    {
                        codeTextBefore += AppendLineEnd(line.Content.Substring(1));
                        codeTextAfter += AppendLineEnd(line.Content.Substring(1));
                    }
                }
                codeTextBefore += "...\r\n";
                codeTextAfter += "...\r\n";
            }

            diffList.Add(codeTextBefore);
            diffList.Add(codeTextAfter);
            return diffList;
        }

        public static string GetCodeFromUrl(string urlPath)
        {
            return "";
        }

        private static string AppendLineEnd(string line)
        {
            if (line.Contains("\r\n"))
            {
                return line;
            }

            if (line.Contains("\r") && !line.Contains("\n"))
            {
                return line + "\n";
            }

            if (line.Contains("\n") && !line.Contains("\r"))
            {
                line = line.Replace("\n", "\r\n");
                return line;
            }

            return line + "\r\n";
        }
    }
}
