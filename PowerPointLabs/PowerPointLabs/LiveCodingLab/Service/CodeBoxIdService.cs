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

        /// <summary>
        /// Add a unique code box id to the hash set.
        /// </summary>
        /// <param name="id">new unique id of the code box</param>
        public static void PopulateCodeBoxIds(int id)
        {
            codeBoxIds.Add(id);
        }

        /// <summary>
        /// Generate a new unique id for a code box
        /// </summary>
        /// <returns>new unique id for code box</returns>
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

        /// <summary>
        /// Clears all ids from the hash set
        /// </summary>
        public static void Clear()
        {
            codeBoxIds.Clear();
        }
    }
}
