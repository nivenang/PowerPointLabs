using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.LiveCodingLab.Service
{
    public class CodeBoxIdService
    {
        private static HashSet<int> codeBoxIds = new HashSet<int>();

        public static void PopulateCodeBoxIds(int id)
        {
            codeBoxIds.Add(id);
        }


        public static int GenerateUniqueId()
        {
            int count = 0;
            do
            {
                count++;
            }
            while (codeBoxIds.Contains(count));

            codeBoxIds.Add(count);
            return count;
        }

        public static void Clear()
        {
            codeBoxIds.Clear();
        }
    }
}
