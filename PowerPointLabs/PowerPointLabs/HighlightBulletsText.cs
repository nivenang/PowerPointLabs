﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class HighlightBulletsText
    {
        public static Color highlightColor = Color.FromArgb(242, 41, 10);
        public static Color defaultColor = Color.FromArgb(0, 0, 0);
        public enum HighlightTextSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightTextSelection userSelection = HighlightTextSelection.kNoneSelected;

        public static void AddHighlightBulletsText()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                //Get shapes to consider for animation
                switch (userSelection)
                {
                    case HighlightTextSelection.kShapeSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        break;
                    case HighlightTextSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();
                        break;
                    case HighlightTextSelection.kNoneSelected:
                        currentSlide.DeleteShapesWithPrefix("PPIndicator");
                        currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
                        selectedShapes = currentSlide.Shapes.Range();
                        break;
                    default:
                        break;
                }

                List<PowerPoint.Shape> shapesToUse = GetShapesToUse(currentSlide, selectedShapes);
                if (currentSlide.Name.Contains("PPTLabsHighlightBulletsSlide"))
                    ProcessExistingHighlightSlide(currentSlide, shapesToUse);

                if (shapesToUse.Count == 0)
                    return;
                currentSlide.Name = "PPTLabsHighlightBulletsSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                bool isFirstShape = IsFirstShape(currentSlide);

                foreach (PowerPoint.Shape sh in shapesToUse)
                {
                    if (!sh.Name.Contains("HighlightTextShape"))
                        sh.Name = "HighlightTextShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                    //Add Font Appear effect for all paragraphs within shape
                    int currentIndex = sequence.Count;
                    sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    var appearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                    //Add Font Disappear effect for all paragraphs within shape
                    currentIndex = sequence.Count;
                    sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    var disappearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                    //Remove effects for paragraphs without bullet points 
                    var markedForRemoval = GetParagraphsToRemove(sh, selectedText);
                    // assert appearEffects.Count == disappearEffects.Count;
                    // assert markedForRemoval.Count <= appearEffects.Count;
                    for (int i = markedForRemoval.Count - 1; i >= 0; --i)
                    {
                        // delete from back.
                        int index = markedForRemoval[i];
                        appearEffects[index].Delete();
                        appearEffects.RemoveAt(index);
                        disappearEffects[index].Delete();
                        disappearEffects.RemoveAt(index);
                    }

                    if (appearEffects.Count == 0)
                    {
                        continue;
                    }

                    RearrangeEffects(appearEffects, disappearEffects);
                    FormatAppearEffects(appearEffects, isFirstShape);
                    FormatDisappearEffects(disappearEffects);
                    isFirstShape = false;
                }

                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddHighlightBulletsText");
                throw;
            }
        }

        /// <summary>
        /// Takes the effects in the sequence in the range [startIndex,endIndex) and puts them into a list in the same order.
        /// </summary>
        private static List<PowerPoint.Effect> AsList(PowerPoint.Sequence sequence, int startIndex, int endIndex)
        {
            var list = new List<PowerPoint.Effect>();
            for (int i = startIndex; i < endIndex; ++i)
            {
                list.Add(sequence[i]);
            }
            return list;
        }

        //Delete existing animations
        private static void ProcessExistingHighlightSlide(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            currentSlide.DeleteShapesWithPrefix("PPIndicator");
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");

            foreach (PowerPoint.Shape tmp in currentSlide.Shapes)
                if (shapesToUse.Contains(tmp))
                    if (userSelection != HighlightTextSelection.kTextSelected)
                        currentSlide.DeleteShapeAnimations(tmp);
        }

        /// <summary>
        /// The add animations creates a text animation for every paragraph in the text box.
        /// But we may not always want all the paragraphs to have animations.
        /// This method marks paragraphs to remove text animations from and returns a list of the indexes of the marked paragraphs.
        /// Indexes returned is in increasing order.
        /// </summary>
        private static List<int> GetParagraphsToRemove(PowerPoint.Shape sh, Office.TextRange2 selectedText)
        {
            var textRange = sh.TextFrame2.TextRange;
            if (userSelection == HighlightTextSelection.kTextSelected)
            {
                return GetUnselectedParagraphs(textRange, selectedText);
            }
            else
            {
                return GetParagraphsWithoutBullets(textRange);
            }
        }

        /// <summary>
        /// If there are bullet points, returns a list of paragraphs without bullet points (marked for removal)
        /// If there are no paragraphs with bullet points at all, then return empty list (mark nothing for removal)
        /// </summary>
        private static List<int> GetParagraphsWithoutBullets(Office.TextRange2 textRange)
        {
            var indexList = new List<int>();
            int index = 0;
            bool hasBulletPoint = false;
            for (int i = 1; i <= textRange.Paragraphs.Count; ++i)
            {
                var paragraph = textRange.Paragraphs[i];
                if (paragraph.Text.Trim().Length == 0)
                {
                    continue;
                }
                if (paragraph.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoFalse)
                {
                    indexList.Add(index);
                }
                else
                {
                    hasBulletPoint = true;
                }
                index++;
            }

            // Return nothing if there is no bullet point at all.
            if (!hasBulletPoint)
            {
                indexList.Clear();
            }
            return indexList;
        }

        /// <summary>
        /// Get a list of unselected paragraphs to mark for removal.
        /// </summary>
        private static List<int> GetUnselectedParagraphs(Office.TextRange2 textRange, Office.TextRange2 selectedText)
        {
            var indexList = new List<int>();
            int index = 0;
            for (int i = 1; i <= textRange.Paragraphs.Count; ++i)
            {
                var paragraph = textRange.Paragraphs[i];
                if (paragraph.Text.Trim().Length == 0)
                {
                    continue;
                }
                if (selectedText.Start + selectedText.Length < paragraph.Start ||
                    selectedText.Start > paragraph.Start + paragraph.Length - 1)
                {
                    indexList.Add(index);
                }
                index++;
            }
            return indexList;
        }


        /// <summary>
        /// Rearranges the appear and disappear effects to be in the correct order for highlight bullets.
        /// Order: [0a] [1a 0d] [2a 1d] [3a 2d] [4a 3d] [4d]
        /// </summary>
        private static void RearrangeEffects(List<PowerPoint.Effect> appearEffects, List<PowerPoint.Effect> disappearEffects)
        {
            // First
            if (appearEffects.Count >= 2)
            {
                appearEffects[1].MoveAfter(appearEffects[0]);
            }

            // Middle
            for (int i = 1; i < appearEffects.Count - 1; ++i)
            {
                disappearEffects[i - 1].MoveAfter(appearEffects[i]);
                appearEffects[i + 1].MoveAfter(disappearEffects[i - 1]);
            }

            // Last
            disappearEffects[appearEffects.Count - 1].MoveAfter(disappearEffects[appearEffects.Count - 1]);
        }

        /// <summary>
        /// Apply formatting and timing to the "appear" effects. (i.e. highlight bullet)
        /// </summary>
        private static void FormatAppearEffects(List<PowerPoint.Effect> appearEffects, bool isFirstShape)
        {
            foreach (var effect in appearEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                // TODO: Orange text bug occurs on this line. effect.EffectParameters.Color2.RGB is not changed for some reason.
                effect.EffectParameters.Color2.RGB = Utils.Graphics.ConvertColorToRgb(highlightColor);
                effect.Timing.Duration = 0.1f;
                effect.Timing.TriggerDelayTime = 0.1f;
            }
            if (!isFirstShape)
            {
                appearEffects.First().Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        /// <summary>
        /// Apply formatting and timing to the "disappear" effects. (i.e. unhighlight bullet)
        /// </summary>
        private static void FormatDisappearEffects(List<PowerPoint.Effect> disappearEffects)
        {
            foreach (var effect in disappearEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                effect.EffectParameters.Color2.RGB = Utils.Graphics.ConvertColorToRgb(defaultColor);
                effect.Timing.Duration = 0.1f;
                effect.Timing.TriggerDelayTime = 0.1f;
            }
            disappearEffects.Last().Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
        }


        private static bool IsFirstShape(PowerPointSlide currentSlide)
        {
            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            bool isFirstShape = true;
            if (sequence.Count != 0)
            {
                isFirstShape = (sequence[sequence.Count].EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor) ? false : true;
                if (sequence[1].EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor)
                    sequence[1].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            return isFirstShape;
        }


        /// <summary>
        /// Get shapes to use for animation.
        /// If user does not select anything: Select shapes which have bullet points
        /// If user selects some shapes: Keep shapes from user selection which have bullet points
        /// If user selects some text: Keep shapes used to store text
        /// </summary>
        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange selectedShapes)
        {
            List<PowerPoint.Shape> shapesToUse = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in selectedShapes)
            {
                if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    shapesToUse.Add(sh);
                }
            }
            return shapesToUse;
        }
    }
}
