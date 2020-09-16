using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

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
                return "Does Not Exist";
            }

            return File.ReadAllText(filePath);

        }

    }
}
