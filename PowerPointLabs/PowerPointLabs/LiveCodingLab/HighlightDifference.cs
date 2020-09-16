using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{
    public partial class LiveCodingLabMain
    {
#pragma warning disable 0618

        internal const int HighlightDifference_MinNoOfShapesRequired = 1;
        internal const string HighlightDifference_FeatureName = "Highlight Difference";
        internal const string HighlightDifference_ShapeSupport = "code box";
        internal static readonly string[] HighlightDifference_ErrorParameters =
{
            HighlightDifference_FeatureName,
            HighlightDifference_MinNoOfShapesRequired.ToString(),
            HighlightDifference_ShapeSupport
        };

        public void HighlightDifferences(PowerPoint.ShapeRange shapeRange)
        {
            try
            {

                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

                if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorHighlightDifferenceWrongSlide,
                                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                    return;
                }

                PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];

                PowerPoint.ShapeRange selectedShapesCurrentSlide = shapeRange;
                PowerPoint.ShapeRange selectedShapesNextSlide = nextSlide.Shapes.Range();

                //Get shapes to consider for animation
                List<PowerPoint.Shape> shapesToUseCurrentSlide = currentSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);
                List<PowerPoint.Shape> shapesToUseNextSlide = nextSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);

                // Check that there exists a "before" code and an "after" code to be animated
                if (shapesToUseCurrentSlide == null || shapesToUseNextSlide == null)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorHighlightDifferenceCodeSnippet,
                                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                    return;
                }

                if (shapesToUseCurrentSlide.Count != 1 || !HasText(shapesToUseCurrentSlide[0]))
                {
                    MessageBox.Show(LiveCodingLabText.ErrorHighlightDifferenceNoSelection,
                                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                    return;
                }

                List<PowerPoint.Shape> shapesToUseNext = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape sh in shapesToUseNextSlide)
                {
                    if (HasText(sh)
                        && sh.TextFrame.TextRange.Paragraphs().Count == shapesToUseCurrentSlide[0].TextFrame.TextRange.Paragraphs().Count)
                    {
                        shapesToUseNext.Add(sh);
                    }
                }

                if (shapesToUseNext.Count < 1)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorHighlightDifferenceCodeSnippet,
                                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                    return;
                }

                PowerPointSlide transitionSlide = currentSlide.Duplicate();
                transitionSlide.Name = "PPTLabsHighlightDifferenceTransitionSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                AddPowerPointLabsIndicator(transitionSlide);

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;

                // Objects that contain the "before" and "after" code to be animated
                PowerPoint.Shape codeShapeBeforeEdit = transitionSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex)[0];
                PowerPoint.Shape codeShapeAfterEdit = transitionSlide.CopyShapeToSlide(shapesToUseNext[0]);
                PowerPoint.TextRange codeTextBeforeEdit = codeShapeBeforeEdit.TextFrame.TextRange;
                PowerPoint.TextRange codeTextAfterEdit = codeShapeAfterEdit.TextFrame.TextRange;

                // Ensure that both pieces of code contain the same number of lines before animating
                if (codeTextBeforeEdit.Paragraphs().Count != codeTextAfterEdit.Paragraphs().Count)
                {
                    return;
                }

                codeShapeAfterEdit.Left = codeShapeBeforeEdit.Left;
                codeShapeAfterEdit.Top = codeShapeBeforeEdit.Top;
                codeShapeAfterEdit.Height = codeShapeBeforeEdit.Height;
                codeShapeAfterEdit.Width = codeShapeBeforeEdit.Width;
                
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

                colourChangeEffectsBefore = DeleteRedundantEffects(markedForRemoval, colourChangeEffectsBefore);

                // Changes colour of text to user-specified colour
                FormatColourChangeEffectsBefore(colourChangeEffectsBefore);
                
                // Creates "appear" effects for "after" code to be transitioned to.
                currentIndex = sequence.Count;
                sequence.AddEffect(
                    codeShapeAfterEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> appearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
               
                appearEffects = DeleteRedundantEffects(markedForRemoval, appearEffects);

                FormatAppearEffectsHighlight(appearEffects);

                // Creates "disappear" effects for "before" to be transitioned away from.
                currentIndex = sequence.Count;
                sequence.AddEffect(
                    codeShapeBeforeEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> disappearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                disappearEffects = DeleteRedundantEffects(markedForRemoval, disappearEffects);

                FormatDisappearEffectsHighlight(disappearEffects);

                // Create colour change effects for the "after" code to highlight code that was changed.
                currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeAfterEdit, 
                    PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, 
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, 
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                List<PowerPoint.Effect> colourChangeEffectsAfter = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                colourChangeEffectsAfter = DeleteRedundantEffects(markedForRemoval, colourChangeEffectsAfter);

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
        private static void FormatAppearEffectsHighlight(List<PowerPoint.Effect> appearEffects)
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
        private static void FormatDisappearEffectsHighlight(List<PowerPoint.Effect> disappearEffects)
        {
            foreach (PowerPoint.Effect effect in disappearEffects)
            {
                effect.Exit = Office.MsoTriState.msoTrue;
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



    }
}
