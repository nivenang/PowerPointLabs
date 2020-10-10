﻿using System.Collections.Generic;

namespace PowerPointLabs.LiveCodingLab
{
    public class ChunkDiff
    {
        public ChunkDiff(string content, int oldStart, int oldLines, int newStart, int newLines)
        {
            Content = content;
            OldStart = oldStart;
            OldLines = oldLines;
            NewStart = newStart;
            NewLines = newLines;
        }

        public ICollection<LineDiff> Changes { get; } = new List<LineDiff>();

        public string Content { get; }

        public int OldStart { get; }

        public int OldLines { get; }

        public int NewStart { get; }

        public int NewLines { get; }
    }
}
