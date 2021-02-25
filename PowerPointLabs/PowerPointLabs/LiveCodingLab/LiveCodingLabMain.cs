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
        private static Shape[] currentSlideShapes;
        private static Shape[] nextSlideShapes;
        private static int[] matchingShapeIDs;
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

        #region Constructor
        public LiveCodingLabMain()
        {
            currentPresentation = PowerPointPresentation.Current;
        }
        #endregion
        public List<int> GetMatchingShapeIDs()
        {
            List<PowerPointSlide> slides = currentPresentation.Slides;
            List<int> matchingShapeIdsToReturn = new List<int>();

            for (int i = 1; i < currentPresentation.SlideCount - 1; i++)
            {
                if (slides[i].HasShapeWithRule(new Regex(@"PPTIndicator.*")) && !slides[i].Hidden)
                {
                    int[] tempMatchingIds = GetMatchingShapeDetails(slides[i - 1], slides[i + 1]);
                    matchingShapeIdsToReturn.AddRange(tempMatchingIds);
                }
            }
            return matchingShapeIdsToReturn;
        }

        private static PowerPointAutoAnimateSlide AddTransitionAnimations(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            PowerPointAutoAnimateSlide addedSlide = currentSlide.CreateAutoAnimateSlide() as PowerPointAutoAnimateSlide;
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);

            foreach (Shape shape in addedSlide.Shapes)
            {
                if (shape.Name.Contains("CodeBox"))
                {
                    shape.Delete();
                }
            }

            addedSlide.MoveMotionAnimation(); // Move shapes with motion animation already added
            addedSlide.PrepareForAutoAnimate();
            if (HasMatchingShapes(currentSlide, nextSlide))
            {
                addedSlide.AddAutoAnimation(currentSlideShapes, nextSlideShapes, matchingShapeIDs);
            }
            return addedSlide;
        }

        private static bool HasMatchingShapes(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            currentSlideShapes = new Shape[currentSlide.Shapes.Count];
            nextSlideShapes = new Shape[currentSlide.Shapes.Count];
            matchingShapeIDs = new int[currentSlide.Shapes.Count];

            int counter = 0;
            PowerPoint.Shape tempMatchingShape = null;
            bool flag = false;

            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                tempMatchingShape = nextSlide.GetShapeWithSameIDAndName(sh);
                if (tempMatchingShape == null)
                {
                    tempMatchingShape = nextSlide.GetShapeWithSameName(sh);
                }

                if (tempMatchingShape != null)
                {
                    currentSlideShapes[counter] = sh;
                    nextSlideShapes[counter] = tempMatchingShape;
                    matchingShapeIDs[counter] = sh.Id;
                    counter++;
                    flag = true;
                }
            }

            return flag;
        }
        
        private int[] GetMatchingShapeDetails(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            Shape[] currentSlideShapesList = new Shape[currentSlide.Shapes.Count];
            Shape[] nextSlideShapesList = new Shape[currentSlide.Shapes.Count];
            int[] matchingShapeIDsList = new int[currentSlide.Shapes.Count];

            int counter = 0;
            PowerPoint.Shape tempMatchingShape = null;

            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                tempMatchingShape = nextSlide.GetShapeWithSameIDAndName(sh);
                if (tempMatchingShape == null)
                {
                    tempMatchingShape = nextSlide.GetShapeWithSameName(sh);
                }

                if (tempMatchingShape != null)
                {
                    currentSlideShapesList[counter] = sh;
                    nextSlideShapesList[counter] = tempMatchingShape;
                    matchingShapeIDsList[counter] = sh.Id;
                    counter++;
                }
            }

            return matchingShapeIDsList;
        }
        /// <summary>
        /// Deletes all redundant effects from the sequence.
        /// </summary>
        /// <param name="markedForRemoval">list of redundant effects to be removed</param>
        /// <param name="effectList">list of effects to be trimmed</param>
        /// <returns>list of effects to be kept</returns>
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
        /// <param name="sequence">sequence of animation effects to appear in</param>
        /// <param name="startIndex">starting index of the effects to be converted to list</param>
        /// <param name="endIndex">ending index of the effects to be converted to list</param>
        /// <returns>list of effects to be made into a list</returns>
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

        /// <summary>
        /// Helper method to add the PowerPointLabs logo to created slides
        /// </summary>
        /// <param name="_slide">slide to insert the logo into</param>
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

