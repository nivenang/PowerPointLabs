using System;
using System.Collections.Generic;
using System.IO;
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

        public static string GetCodeFromUrl(string urlPath)
        {
            return "";
        }

    }
}
