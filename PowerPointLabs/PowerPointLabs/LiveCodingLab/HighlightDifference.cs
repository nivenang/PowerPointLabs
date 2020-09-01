using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{
    class HighlightDifference
    {
#pragma warning disable 0618

        public static bool IsHighlightDifferenceEnabled { get; set; } = true;

        public static void HighlightDifferences()
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                //Get shapes to consider for animation
                List<PowerPoint.Shape> shapesToUse = GetShapesToUse(currentSlide, selectedShapes);

                // Delete all existing animations
                if (currentSlide.Name.Contains("PPTLabsHighlightDifferenceSlide"))
                {
                    ProcessExistingHighlightSlide(currentSlide, shapesToUse);
                }

                // Check that there exists a "before" code and an "after" code to be animated
                if (shapesToUse == null || shapesToUse.Count != 2)
                {
                    return;
                }

                // Set a flag that shows that there are existing animations in the current slide
                currentSlide.Name = "PPTLabsHighlightDifferenceSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;

                // Objects that contain the "before" and "after" code to be animated
                PowerPoint.Shape codeShapeBeforeEdit = shapesToUse[0];
                PowerPoint.Shape codeShapeAfterEdit = shapesToUse[1];
                PowerPoint.TextRange codeTextBeforeEdit = codeShapeBeforeEdit.TextFrame.TextRange;
                PowerPoint.TextRange codeTextAfterEdit = codeShapeAfterEdit.TextFrame.TextRange;

                // Ensure that both pieces of code contain the same number of lines before animating
                if (codeTextBeforeEdit.Paragraphs().Count != codeTextAfterEdit.Paragraphs().Count)
                {
                    return;
                }

                // Create a duplicate code over the before to simulate animation between a "before" and "after" code
                PowerPoint.Shape codeShapeIntermediateEdit = codeShapeAfterEdit.Duplicate()[1];
                codeShapeIntermediateEdit.Left = codeShapeBeforeEdit.Left;
                codeShapeIntermediateEdit.Top = codeShapeBeforeEdit.Top;
                codeShapeIntermediateEdit.Height = codeShapeBeforeEdit.Height;
                codeShapeIntermediateEdit.Width = codeShapeBeforeEdit.Width;
                currentSlide.RemoveAnimationsForShape(codeShapeIntermediateEdit);

                // Remove the "after" code for alignment with "before" code using the previously created duplicate.
                PowerPoint.Effect codeShapeAfterEditDisappear = sequence.AddEffect(codeShapeAfterEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                codeShapeAfterEditDisappear.Exit = Office.MsoTriState.msoTrue;
                codeShapeAfterEditDisappear.Timing.Duration = 0;
                
                // Add Colour change effect for lines of code to be changed.
                int currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeBeforeEdit, 
                    PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, 
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, 
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> colourChangeEffectsBefore = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                // Removes colour change effect from all lines of code that are not changed.
                List<int> markedForRemoval = new List<int>();
                int effectCount = 0;

                for (int paragraphCount = 0; paragraphCount < codeTextBeforeEdit.Paragraphs().Count; paragraphCount++)
                {
                    if (codeTextBeforeEdit.Paragraphs(paragraphCount+1).TrimText().Text == "")
                    {
                        continue;
                    }

                    if (codeTextBeforeEdit.Paragraphs(paragraphCount+1).TrimText().Text == codeTextAfterEdit.Paragraphs(paragraphCount+1).TrimText().Text)
                    {
                        markedForRemoval.Add(effectCount);
                    }

                    effectCount++;
                }

                for (int i = markedForRemoval.Count - 1; i >= 0; --i)
                {
                    // delete redundant colour change effects from back.
                    int index = markedForRemoval[i];
                    colourChangeEffectsBefore[index].Delete();
                    colourChangeEffectsBefore.RemoveAt(index);
                }

                // Changes colour of text to user-specified colour
                FormatColourChangeEffectsBefore(colourChangeEffectsBefore);
                
                // Creates "appear" effects for "after" code to be transitioned to.
                currentIndex = sequence.Count;
                sequence.AddEffect(
                    codeShapeIntermediateEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> appearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                for (int i = markedForRemoval.Count - 1; i >= 0; --i)
                {
                    // delete redundant appear effects from back.
                    int index = markedForRemoval[i];
                    appearEffects[index].Delete();
                    appearEffects.RemoveAt(index);
                }

                FormatAppearEffects(appearEffects);

                // Creates "disappear" effects for "before" to be transitioned away from.
                currentIndex = sequence.Count;
                sequence.AddEffect(
                    codeShapeBeforeEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> disappearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                foreach (PowerPoint.Effect disappearEffect in disappearEffects)
                {
                    disappearEffect.Exit = Office.MsoTriState.msoTrue;
                }

                for (int i = markedForRemoval.Count - 1; i >= 0; --i)
                {
                    // delete redundant "disappear" effects from back.
                    int index = markedForRemoval[i];
                    disappearEffects[index].Delete();
                    disappearEffects.RemoveAt(index);
                }

                FormatDisappearEffects(disappearEffects);

                // Create colour change effects for the "after" code to highlight code that was changed.
                currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeIntermediateEdit, 
                    PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, 
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, 
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                List<PowerPoint.Effect> colourChangeEffectsAfter = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                for (int i = markedForRemoval.Count - 1; i >= 0; --i)
                {
                    // delete redundant colour change effects from back.
                    int index = markedForRemoval[i];
                    colourChangeEffectsAfter[index].Delete();
                    colourChangeEffectsAfter.RemoveAt(index);
                }

                // Changes colour of text to user-specified colour
                FormatColourChangeEffectsAfter(colourChangeEffectsAfter);
                
                // Re-orders the effects to create a full highlight difference animation
                RearrangeEffects(colourChangeEffectsBefore, appearEffects, disappearEffects, colourChangeEffectsAfter);

                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HighlightDifferences");
                throw;
            }
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

        // Delete existing animations
        private static void ProcessExistingHighlightSlide(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            currentSlide.DeleteIndicator();

            foreach (PowerPoint.Shape tmp in currentSlide.Shapes)
            {
                if (shapesToUse.Contains(tmp))
                {
                    currentSlide.DeleteShapeAnimations(tmp);
                }
            }
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
        /// Apply colour change and timing to the lines of code that is going to disappear (i.e. code to be changed from).
        /// </summary>
        private static void FormatColourChangeEffectsBefore(List<PowerPoint.Effect> colourChangeEffects)
        {
            foreach (PowerPoint.Effect effect in colourChangeEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                // TODO: Orange text bug occurs on this line. effect.EffectParameters.Color2.RGB is not changed for some reason.
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
                effect.Timing.Duration = 0.1f;
                effect.Timing.TriggerDelayTime = 0.1f;
            }
        }

        /// <summary>
        /// Apply formatting and timing to the "appear" effects (i.e. new code to be changed to).
        /// </summary>
        private static void FormatAppearEffects(List<PowerPoint.Effect> appearEffects)
        {
            foreach (PowerPoint.Effect effect in appearEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                effect.Timing.Duration = 0.1f;
                effect.Timing.TriggerDelayTime = 0.1f;
            }
        }

        /// <summary>
        /// Apply colour change and timing to the lines of code that is going to appear (i.e. code to be changed to).
        /// </summary>
        private static void FormatColourChangeEffectsAfter(List<PowerPoint.Effect> colourChangeEffects)
        {
            foreach (PowerPoint.Effect effect in colourChangeEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                // TODO: Orange text bug occurs on this line. effect.EffectParameters.Color2.RGB is not changed for some reason.
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
                effect.Timing.Duration = 0;
            }
        }

        /// <summary>
        /// Apply formatting and timing to the "disappear" effects. (i.e. old code to be changed from)
        /// </summary>
        private static void FormatDisappearEffects(List<PowerPoint.Effect> disappearEffects)
        {
            foreach (PowerPoint.Effect effect in disappearEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                effect.Timing.Duration = 0;
            }
        }

        /// <summary>
        /// Rearranges the colour change, appear and disappear effects to be in the correct order for highlight differences.
        /// Order: [0colour change 0disappear 0appear 0colour change] [1cc 1d 1a 1cc] ...
        /// </summary>
        private static void RearrangeEffects(List<PowerPoint.Effect> colourChangeEffectsBefore, List<PowerPoint.Effect> appearEffects, List<PowerPoint.Effect> disappearEffects, List<PowerPoint.Effect> colourChangeEffectsAfter)
        {
            if (colourChangeEffectsBefore.Count <= 0 || appearEffects.Count <= 0 || disappearEffects.Count <= 0 || colourChangeEffectsAfter.Count <= 0)
            {
                return;
            }

            for (int i = 0; i < colourChangeEffectsBefore.Count; i++)
            {
                if (i >= 1)
                {
                    colourChangeEffectsBefore[i].MoveAfter(colourChangeEffectsAfter[i - 1]);
                }
                disappearEffects[i].MoveAfter(colourChangeEffectsBefore[i]);
                appearEffects[i].MoveAfter(disappearEffects[i]);
                colourChangeEffectsAfter[i].MoveAfter(appearEffects[i]);
            }
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
