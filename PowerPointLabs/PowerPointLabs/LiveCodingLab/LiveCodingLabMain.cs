using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
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
    public partial class LiveCodingLabMain
    {
        public LiveCodingLabMain()
        {

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
        private static bool HasText(PowerPoint.Shape shape)
        {
            return shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                   shape.TextFrame2.HasText == Office.MsoTriState.msoTrue;

        }
    }
}

