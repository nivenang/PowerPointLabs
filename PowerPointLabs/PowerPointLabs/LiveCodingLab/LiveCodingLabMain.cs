using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Drawing = System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.LiveCodingLab
{
#pragma warning disable 0618
    public partial class LiveCodingLabMain
    {
        PowerPointPresentation currentPresentation;
        enum DiffType
        {
            Add,
            Delete,
            Normal
        }

        enum WordDiffType
        {
            DeleteEqual,
            DeleteAdd,
            AddEqual,
            Equal
        }

        public LiveCodingLabMain()
        {
            currentPresentation = PowerPointPresentation.Current;
        }

        /// <summary>
        /// Deletes all redundant effects from the sequence.
        /// </summary>
        private static List<PowerPoint.Effect> DeleteRedundantEffects(List<int> markedForRemoval, List<PowerPoint.Effect> effectList)
        {
            for (int i = markedForRemoval.Count - 1; i >= 0; --i)
            {
                // delete redundant colour change effects from back.
                int index = markedForRemoval[i];
                effectList[index].Delete();
                effectList.RemoveAt(index);
            }
            return effectList;
        }

        /// <summary>
        /// Takes the effects in the sequence in the range [startIndex,endIndex) and puts them into a list in the same order.
        /// </summary>
        private static List<PowerPoint.Effect> AsList(PowerPoint.Sequence sequence, int startIndex, int endIndex)
        {
            List<PowerPoint.Effect> list = new List<PowerPoint.Effect>();
            for (int i = startIndex; i < endIndex; ++i)
            {
                list.Add(sequence[i]);
            }
            return list;
        }

        /// <summary>
        /// Get shapes to use for animation.
        /// If user does not select anything: Select shapes which have bullet points
        /// If user selects some shapes: Keep shapes from user selection which have bullet points
        /// If user selects some text: Keep shapes used to store text
        /// </summary>
        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange selectedShapes)
        {
            return selectedShapes.Cast<PowerPoint.Shape>()
                                .Where(HasText)
                                .ToList();
        }

        /// <summary>
        /// Returns true iff shape has a text frame.
        /// </summary>
        private static bool HasText(Shape shape)
        {
            return shape.HasTextFrame == MsoTriState.msoTrue &&
                   shape.TextFrame2.HasText == MsoTriState.msoTrue;
        }

        private void AddPowerPointLabsIndicator(PowerPointSlide _slide)
        {
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            Shape indicatorShape = _slide.Shapes.AddPicture(tempFileName, MsoTriState.msoFalse, MsoTriState.msoTrue, PowerPointPresentation.Current.SlideWidth - 120, 0, 120, 84);

            indicatorShape.Left = PowerPointPresentation.Current.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = PowerPointSlide.PptLabsIndicatorShapeName + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            ShapeUtil.MakeShapeViewTimeInvisible(indicatorShape, _slide);
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

