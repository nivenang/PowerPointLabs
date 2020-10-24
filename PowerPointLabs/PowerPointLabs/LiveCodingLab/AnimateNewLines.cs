using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Views;
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
        internal const int AnimateNewLines_MinNoOfShapesRequired = 1;
        internal const string AnimateNewLines_FeatureName = "Animate New Lines";
        internal const string AnimateNewLines_ShapeSupport = "code box";
        internal static readonly string[] AnimateNewLines_ErrorParameters =
        {
            AnimateNewLines_FeatureName,
            AnimateNewLines_MinNoOfShapesRequired.ToString(),
            AnimateNewLines_ShapeSupport
        };

        private static float fontScale = 4.5f;

        public void AnimateNewLines(List<CodeBoxPaneItem> listCodeBox)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

                if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesWrongSlide,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];

                //Get shapes to consider for animation
                CodeBoxPaneItem currentSlideCodeBox = listCodeBox[0];
                CodeBoxPaneItem nextSlideCodeBox = listCodeBox[1];

                PowerPoint.Shape currentSlideShape = currentSlideCodeBox.CodeBox.Shape;

                // Check that there exists a "before" code and an "after" code to be animated
                if (currentSlideCodeBox.CodeBox.Shape == null || nextSlideCodeBox.CodeBox.Shape == null)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                // Retrieves all possible matching code snippets from the next slide
                if (nextSlideCodeBox.CodeBox.Shape.TextFrame.TextRange.Lines().Count <= currentSlideCodeBox.CodeBox.Shape.TextFrame.TextRange.Lines().Count)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesWrongCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                nextSlideCodeBox.CodeBox.Shape.Left = currentSlideCodeBox.CodeBox.Shape.Left;
                nextSlideCodeBox.CodeBox.Shape.Top = currentSlideCodeBox.CodeBox.Shape.Top;
                nextSlideCodeBox.CodeBox.Shape.Width = currentSlideCodeBox.CodeBox.Shape.Width;
                nextSlideCodeBox.CodeBox.Shape.Height = currentSlideCodeBox.CodeBox.Shape.Height;

                // Creates a new animation slide between the before and after code
                PowerPointSlide transitionSlide = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlide.Index + 1);
                transitionSlide.Name = LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                AddPowerPointLabsIndicator(transitionSlide);

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;

                // Objects that contain the "before" and "after" code to be animated
                PowerPoint.Shape codeShapeBeforeEdit = transitionSlide.CopyShapeToSlide(currentSlideCodeBox.CodeBox.Shape);
                PowerPoint.Shape codeShapeAfterEdit = transitionSlide.CopyShapeToSlide(nextSlideCodeBox.CodeBox.Shape);
                PowerPoint.TextRange codeTextBeforeEdit = codeShapeBeforeEdit.TextFrame.TextRange;
                PowerPoint.TextRange codeTextAfterEdit = codeShapeAfterEdit.TextFrame.TextRange;
                

                // Stores the font size of the code snippet for animation scaling
                float fontSize = codeTextBeforeEdit.Font.Size;

                // Aligns the after code with the before code for animation
                codeShapeAfterEdit.Left = codeShapeBeforeEdit.Left;
                codeShapeAfterEdit.Top = codeShapeBeforeEdit.Top;
                codeShapeAfterEdit.Height = codeShapeBeforeEdit.Height;
                codeShapeAfterEdit.Width = codeShapeBeforeEdit.Width;

                // Lists that stores the mapping of each line of the before code to the line number in the after code 
                // e.g. Line 3 of the before code might map to Line 5 of the after code with the added lines
                List<int> codeBeforeMapping = new List<int>();
                List<int> codeAfterMapping = new List<int>();

                // Lists that store the redundant effects
                List<int> markedForAppearRemoval = new List<int>();
                List<int> markedForDisappearRemoval = new List<int>();
                
                // Pointers to keep track of the paragraph counts of the after code and the effects count respectively
                int paragraphCountAfter = 1;
                int effectCount = 0;

                // Populates the mapping for before code to the after code
                // and keeps track of the lines to remove redundant effects from
                for (int paragraphCount = 1; paragraphCount <= codeTextBeforeEdit.Paragraphs().Count; paragraphCount++)
                {
                    // Skip the line if it is empty
                    if (codeTextBeforeEdit.Paragraphs(paragraphCount).TrimText().Text == "")
                    {
                        paragraphCountAfter++;
                        continue;
                    }

                    // If the lines of code are the same between the before and after code, store the mapping in the lists and remove redundant effects from it
                    if (codeTextBeforeEdit.Paragraphs(paragraphCount).TrimText().Text == codeTextAfterEdit.Paragraphs(paragraphCountAfter).TrimText().Text)
                    {
                        codeBeforeMapping.Add(paragraphCount);
                        codeAfterMapping.Add(paragraphCountAfter);
                        markedForAppearRemoval.Add(effectCount);
                        effectCount++;
                    }
                    // If the lines of code are not the same, keep traversing down the after code until there is a match and then store the mapping.
                    else
                    {
                        while (codeTextBeforeEdit.Paragraphs(paragraphCount).TrimText().Text != codeTextAfterEdit.Paragraphs(paragraphCountAfter).TrimText().Text)
                        {
                            if (paragraphCountAfter + 1 > codeTextAfterEdit.Paragraphs().Count)
                            {
                                MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesWrongCodeSnippet,
                                                LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                                transitionSlide.Delete();
                                nextSlideCodeBox.CodeBox.Slide = nextSlide;
                                return;
                            }
                            paragraphCountAfter++;
                            markedForDisappearRemoval.Add(effectCount);
                            effectCount++;
                        }
                        codeBeforeMapping.Add(paragraphCount);
                        codeAfterMapping.Add(paragraphCountAfter);
                        markedForAppearRemoval.Add(effectCount);
                        effectCount++;
                    }
                    paragraphCountAfter++;
                }

                // Creates disappear effects to remove lines that are similar between both codes
                int currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeAfterEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> disappearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                disappearEffects = DeleteRedundantEffects(markedForDisappearRemoval, disappearEffects);
                FormatDisappearEffects(disappearEffects);

                // Creates appear effects for new lines of code to be inserted
                currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeAfterEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                List<PowerPoint.Effect> appearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                appearEffects = DeleteRedundantEffects(markedForAppearRemoval, appearEffects);
                FormatAppearEffects(appearEffects);

                // Goes through before and after code mapping to obtain number of lines to move new lines down by
                List<int> markedForRemoval = new List<int>();
                List<int> multiplierQueue = new List<int>();
                int markedForRemovalPointer = 0;

                for (int i = 0; i < codeBeforeMapping.Count; i++)
                {
                    if (codeBeforeMapping[i] == codeAfterMapping[i])
                    {
                        markedForRemoval.Add(i);
                        markedForRemovalPointer = i;
                    }
                    else if (multiplierQueue.Contains(codeAfterMapping[i] - codeBeforeMapping[i]))
                    {
                        continue;
                    }
                    else
                    {
                        multiplierQueue.Add(codeAfterMapping[i] - codeBeforeMapping[i]);
                    }
                }

                // Creates move down effects for lines below the newly inserted line
                int previousMultiplier = 0;
                for (int j = 0; j < multiplierQueue.Count; j++) 
                {
                    // Creates move down effects for lines below the newly inserted line
                    currentIndex = sequence.Count;
                    sequence.AddEffect(codeShapeBeforeEdit,
                        PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    List<PowerPoint.Effect> moveDownEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                    moveDownEffects = DeleteRedundantEffects(markedForRemoval, moveDownEffects);
                    
                    // Formats move down effects by scaling move down distance with fontsize and number of new lines inserted
                    for (int effect = 0; effect < moveDownEffects.Count; effect++)
                    {
                        PowerPoint.AnimationBehavior behaviour = moveDownEffects[effect].Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                        behaviour.MotionEffect.FromX = 0;
                        behaviour.MotionEffect.FromY = (fontSize / fontScale) * previousMultiplier;
                        behaviour.MotionEffect.ToX = 0;
                        behaviour.MotionEffect.ToY = (fontSize / fontScale) * (multiplierQueue[j]);
                        moveDownEffects[effect].Timing.Duration = 0.5f;
                        if (effect == 0)
                        {
                            moveDownEffects[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                        }
                        else
                        {
                            moveDownEffects[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }

                    }

                    // Rearrange appear effects to make the new line appear right after the move down effect
                    for (int appear = previousMultiplier; appear < multiplierQueue[j]; appear++)
                    {
                        if (appear == previousMultiplier)
                        {
                            appearEffects[appear].MoveAfter(moveDownEffects[moveDownEffects.Count - 1]);
                        }
                        else
                        {
                            appearEffects[appear].MoveAfter(appearEffects[appear - 1]);
                        }
                    }

                    previousMultiplier = multiplierQueue[j];

                    // Add the lines with the newly created effects to the redundant list 
                    // so that effects created in the next iteration does not apply to the those lines
                    for (int k = markedForRemovalPointer + 1; k < codeBeforeMapping.Count; k++)
                    {
                        if (codeAfterMapping[k] - codeBeforeMapping[k] == previousMultiplier)
                        {
                            markedForRemoval.Add(k);
                            markedForRemovalPointer = k;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                nextSlideCodeBox.CodeBox.Slide = nextSlide;
                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateNewLines");
                throw;
            }
        }

        /// <summary>
        /// Apply formatting and timing to the "appear" effects (i.e. new code to be changed to).
        /// </summary>
        private static void FormatAppearEffects(List<PowerPoint.Effect> appearEffects)
        {
            foreach (PowerPoint.Effect effect in appearEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
                effect.Timing.Duration = 0.5f;
            }
        }

        /// <summary>
        /// Apply formatting and timing to the "disappear" effects (i.e. repetitive code).
        /// </summary>
        private static void FormatDisappearEffects(List<PowerPoint.Effect> disappearEffects)
        {
            foreach (PowerPoint.Effect effect in disappearEffects)
            {
                effect.Exit = Office.MsoTriState.msoTrue;
                effect.Timing.Duration = 0;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
        }

        /// <summary>
        /// Apply colour change and timing to the lines of code that is going to appear (i.e. code to be changed to).
        /// </summary>
        private static void FormatColourChangeEffects(List<PowerPoint.Effect> colourChangeEffects)
        {
            foreach (PowerPoint.Effect effect in colourChangeEffects)
            {
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                // TODO: Orange text bug occurs on this line. effect.EffectParameters.Color2.RGB is not changed for some reason.
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
                effect.Timing.Duration = 0;
            }
        }
    }
}
