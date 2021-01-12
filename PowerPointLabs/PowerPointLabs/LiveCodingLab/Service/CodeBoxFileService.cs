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
        /// <summary>
        /// Retrieves code from a file and converts to a string for further processing
        /// </summary>
        /// <param name="filePath">filePath of the file containing the code</param>
        /// <returns>parsed code from file in string format</returns>
        public static string GetCodeFromFile(string filePath)
        {
            // Check if file exists, inform user if file does not exist
            if (!File.Exists(filePath))
            {
                MessageBox.Show(LiveCodingLabText.ErrorInvalidFileName,
                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                return "";
            }

            return File.ReadAllText(filePath);

        }

        /// <summary>
        /// Parses a diff into a list of changes by each diff block
        /// </summary>
        /// <param name="diffPath">file path of the diff file containing the code</param>
        /// <returns>list of differences for each block</returns>
        public static List<FileDiff> ParseDiff(string diffPath)
        {
            // Retrieves the code from the diff file
            string diffInput = GetCodeFromFile(diffPath);

            if (diffInput == "")
            {
                return null;
            }

            // Parses the diff file into diffs by block
            List<FileDiff> diffList = Diff.Parse(diffInput, Environment.NewLine).ToList();
            return diffList;
        }

        /// <summary>
        /// Converts a diff block into a "before" and "after" code
        /// </summary>
        /// <param name="diffFile">One diff block containing code of one block</param>
        /// <returns>a list containing strings of the "before" and "after" code snippets</returns>
        public static List<string> ConvertFileDiffToString(FileDiff diffFile)
        {
            List<ChunkDiff> diffChunks = diffFile.Chunks.ToList();
            List<string> diffList = new List<string>();
            string codeTextBefore = "";
            string codeTextAfter = "";

            // Convert the diff block into "before" and "after" code
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

        /// <summary>
        /// Helper method to append line ends to each line
        /// </summary>
        /// <param name="line">Line to append the line end to</param>
        /// <returns>line containing the appended line end</returns>
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
