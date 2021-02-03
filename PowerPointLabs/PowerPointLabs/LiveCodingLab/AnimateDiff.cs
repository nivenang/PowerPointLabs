using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Utility;
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

        private void AnimateDiff (List<CodeBoxPaneItem> listCodeBox, FileDiff diff, bool isBlockDiff)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

                CodeBoxPaneItem diffCodeBoxBefore = listCodeBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = listCodeBox[1];

                List<ChunkDiff> diffChunks = diff.Chunks.ToList();
                Dictionary<int, DiffType> fullDiff = new Dictionary<int, DiffType>();

                List<int> markedForDisappear = new List<int>();
                List<int> markedForAppear = new List<int>();
                int beforeCounter = 0;
                int afterCounter = 0;
                int lineCounter = 0;

                // Get the diff type for each code line
                foreach (ChunkDiff chunk in diffChunks)
                {
                    List<LineDiff> diffLines = chunk.Changes.ToList();
                    foreach (LineDiff line in diffLines)
                    {

                        if (line.Add)
                        {
                            markedForAppear.Add(afterCounter);
                            afterCounter++;
                            fullDiff.Add(lineCounter, DiffType.Add);
                        }
                        else if (line.Delete)
                        {
                            markedForDisappear.Add(beforeCounter);
                            beforeCounter++;
                            fullDiff.Add(lineCounter, DiffType.Delete);
                        }
                        else
                        {
                            fullDiff.Add(lineCounter, DiffType.Normal);
                            beforeCounter++;
                            afterCounter++;
                        }
                        lineCounter++;
                    }
                    fullDiff.Add(lineCounter, DiffType.Normal);
                    lineCounter++;
                    beforeCounter++;
                    afterCounter++;
                }

                // Creates a new animation slide between the before and after code
                PowerPointSlide transitionSlide = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlide.Index + 1);
                
                if (isBlockDiff)
                {
                    transitionSlide.Name = LiveCodingLabText.AnimateBlockDiffIdentifier + LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                }
                else
                {
                    transitionSlide.Name = LiveCodingLabText.AnimateLineDiffIdentifier + LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                }
                AddPowerPointLabsIndicator(transitionSlide);

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;

                PowerPoint.Shape codeShapeBeforeEdit = transitionSlide.CopyShapeToSlide(diffCodeBoxBefore.CodeBox.Shape);
                PowerPoint.Shape codeShapeAfterEdit = transitionSlide.CopyShapeToSlide(diffCodeBoxAfter.CodeBox.Shape);
                codeShapeBeforeEdit.Name = LiveCodingLabText.TransitionTextIdentifier + "Before" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                codeShapeAfterEdit.Name = LiveCodingLabText.TransitionTextIdentifier + "After" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.TextRange codeTextBeforeEdit = codeShapeBeforeEdit.TextFrame.TextRange;
                PowerPoint.TextRange codeTextAfterEdit = codeShapeAfterEdit.TextFrame.TextRange;
                
                // Stores the font size of the code snippet for animation scaling
                float fontSize = codeTextBeforeEdit.Font.Size;

                // Aligns the after code with the before code for animation
                codeShapeAfterEdit.Left = codeShapeBeforeEdit.Left;
                codeShapeAfterEdit.Top = codeShapeBeforeEdit.Top;
                codeShapeAfterEdit.Height = codeShapeBeforeEdit.Height;
                codeShapeAfterEdit.Width = codeShapeBeforeEdit.Width;
                
                // Creates disappear effects to remove lines that are similar between both codes
                int currentIndex = sequence.Count;
                sequence.AddEffect(codeShapeAfterEdit,
                    PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                List<PowerPoint.Effect> disappearEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
               
                FormatDisappearEffects(disappearEffects);

                Dictionary<int, int> beforeLineToEffectLine = new Dictionary<int, int>();
                Dictionary<int, int> afterLineToEffectLine = new Dictionary<int, int>();
                int effectCount = 0;

                // Create a mapping from line number to index of effect in "before" code effect list
                for (int i = 0; i < codeTextBeforeEdit.Paragraphs().Count; i++)
                {
                    if (codeTextBeforeEdit.Paragraphs(i+1).TrimText().Text == "")
                    {
                        beforeLineToEffectLine.Add(i, effectCount);
                        continue;
                    }
                    beforeLineToEffectLine.Add(i, effectCount);
                    effectCount++;
                }

                effectCount = 0;

                // Create a mapping from line number to index of effect in "after" code effect list
                for (int i = 0; i < codeTextAfterEdit.Paragraphs().Count; i++)
                {
                    if (codeTextAfterEdit.Paragraphs(i + 1).TrimText().Text == "")
                    {
                        afterLineToEffectLine.Add(i, effectCount);
                        continue;
                    }
                    afterLineToEffectLine.Add(i, effectCount);
                    effectCount++;
                }

                int beforeCount = 0;
                int afterCount = 0;
                int currentMultiplier = 0;
                int lineCount = 0;
                List<PowerPoint.Effect> disappearHighlightEffects = new List<PowerPoint.Effect>();
                List<PowerPoint.Effect> appearHighlightEffects = new List<PowerPoint.Effect>();
                List<PowerPoint.Effect> intermediateAppearEffects = new List<PowerPoint.Effect>();
                List<PowerPoint.Effect> intermediateDisappearEffects = new List<PowerPoint.Effect>();

                // Keep generating animations for code lines while there are still code lines to be animated
                while (beforeCount < codeTextBeforeEdit.Paragraphs().Count && afterCount < codeTextAfterEdit.Paragraphs().Count)
                {
                    // Case 1: Code line is to be deleted
                    if (fullDiff[lineCount] == DiffType.Delete)
                    {
                        // Case 1a: Deletion line is a newline, animate only the subsequent move up effects, no deletion animation necessary
                        if (codeTextBeforeEdit.Paragraphs(beforeCount + 1).TrimText().Text == "")
                        {
                            // Check that there is no addition line after the line to be deleted
                            // If there is an addition line after the line to be deleted, do not animate the shifting up,
                            // simply animate the appearance of addition line to minimise shifting effects
                            if (lineCount + 1 >= fullDiff.Count || (lineCount + 1 < fullDiff.Count && fullDiff[lineCount + 1] != DiffType.Add))
                            {
                                currentIndex = sequence.Count;
                                sequence.AddEffect(codeShapeBeforeEdit,
                                    PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                                moveUpEffects = FormatMoveUpWhitespaceEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);

                                intermediateDisappearEffects.AddRange(moveUpEffects);

                                currentMultiplier--;
                            }
                            beforeCount++;
                            lineCount++;
                            continue;
                        }

                        // Case 1b: Deletion line contains text, animate deletion of line and subsequent shifting up of all code below it.
                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeEffectsBefore = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        colourChangeEffectsBefore = FormatDisappearColourChangeEffects(beforeLineToEffectLine[beforeCount], colourChangeEffectsBefore);

                        disappearHighlightEffects.AddRange(colourChangeEffectsBefore);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> deleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        deleteEffects = FormatDeleteEffects(beforeLineToEffectLine[beforeCount], deleteEffects);

                        intermediateDisappearEffects.AddRange(deleteEffects);

                        // Check that there is no addition line after the line to be deleted
                        // If there is an addition line after the line to be deleted, do not animate the shifting up,
                        // simply animate the appearance of addition line to minimise shifting effects
                        if (lineCount + 1 >= fullDiff.Count || (lineCount + 1 < fullDiff.Count && fullDiff[lineCount + 1] != DiffType.Add))
                        {
                            currentIndex = sequence.Count;
                            sequence.AddEffect(codeShapeBeforeEdit,
                                PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                                PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                            moveUpEffects = FormatMoveUpEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);

                            intermediateDisappearEffects.AddRange(moveUpEffects);

                            currentMultiplier--;
                        }

                        beforeCount++;
                        lineCount++;
                    }
                    // Case 2: Code line is to be inserted/added
                    else if (fullDiff[lineCount] == DiffType.Add)
                    {
                        // Case 2a: Addition line is a newline, animate only the subsequent move down effects, no addition animation necessary
                        if (codeTextAfterEdit.Paragraphs(afterCount + 1).TrimText().Text == "")
                        {
                            if (lineCount == 0 || (lineCount - 1 >= 0 && fullDiff[lineCount - 1] != DiffType.Delete))
                            {
                                currentIndex = sequence.Count;
                                sequence.AddEffect(codeShapeBeforeEdit,
                                    PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                List<PowerPoint.Effect> moveDownEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                                moveDownEffects = FormatMoveDownWhitespaceEffects(beforeLineToEffectLine[beforeCount], moveDownEffects, currentMultiplier, fontSize);

                                intermediateAppearEffects.AddRange(moveDownEffects);

                                currentMultiplier++;
                            }
                            afterCount++;
                            lineCount++;
                            continue;
                        }

                        // Case 2b: Addition line has text, animate both move down effects and addition line

                        // Check if there is a deletion line before the addition line
                        // If there is no deletion line, animate move down effects to accommodate new addition line
                        if (lineCount == 0 || (lineCount - 1 >= 0 && fullDiff[lineCount - 1] != DiffType.Delete))
                        {
                            currentIndex = sequence.Count;
                            sequence.AddEffect(codeShapeBeforeEdit,
                                PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                                PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            List<PowerPoint.Effect> moveDownEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                            moveDownEffects = FormatMoveDownEffects(beforeLineToEffectLine[beforeCount], moveDownEffects, currentMultiplier, fontSize);

                            intermediateAppearEffects.AddRange(moveDownEffects);

                            currentMultiplier++;
                        }

                        // Animate the appearance of the new line
                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> insertEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        insertEffects = FormatInsertEffects(afterLineToEffectLine[afterCount], insertEffects);

                        intermediateAppearEffects.AddRange(insertEffects);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeEffectsAfter = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        colourChangeEffectsAfter = FormatAppearColourChangeEffects(afterLineToEffectLine[afterCount], colourChangeEffectsAfter);

                        appearHighlightEffects.AddRange(colourChangeEffectsAfter);

                        afterCount++;
                        lineCount++;
                    }
                    // Case 3: Code line exists in both "before" and "after" code, so no animation is created
                    else
                    {
                        beforeCount++;
                        afterCount++;
                        lineCount++;
                    }
                }

                // Rearrange effects to change code from "before" to "after" block by block
                if (isBlockDiff)
                {
                    RearrangeBlockDiffEffects(disappearHighlightEffects, disappearEffects[disappearEffects.Count - 1], PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    RearrangeBlockDiffEffects(intermediateDisappearEffects, disappearHighlightEffects[disappearHighlightEffects.Count - 1], PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    RearrangeBlockDiffEffects(intermediateAppearEffects, intermediateDisappearEffects[intermediateDisappearEffects.Count - 1], PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    RearrangeBlockDiffEffects(appearHighlightEffects, intermediateAppearEffects[intermediateAppearEffects.Count - 1], PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateDiff");
                throw;
            }
        }

        /// <summary>
        /// Removes all redundant delete effects and modifies effect animation to simulate a delete animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatDeleteEffects(int lineToKeep, List<PowerPoint.Effect> effectList)
        {
            for (int i = effectList.Count - 1; i >= 0; --i)
            {
                // delete redundant colour change effects from back.
                if (i != lineToKeep)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }

            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionRight;
                effect.Exit = Office.MsoTriState.msoTrue;
                effect.Timing.Duration = 0.7f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant move up effects and modifies effect animation to simulate a move up animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <param name="currentMultiplier">multiplier for the distance that the effect must travel</param>
        /// <param name="fontSize">size of font to scale the effects movement distance</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatMoveUpEffects(int lineToKeep, List<PowerPoint.Effect> effectList, int currentMultiplier, float fontSize)
        {
            for (int i = lineToKeep; i >= 0; --i)
            {
                effectList[i].Delete();
                effectList.RemoveAt(i);
            }

            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.5f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                PowerPoint.AnimationBehavior behaviour = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.FromY = (fontSize / fontScale) * currentMultiplier;
                behaviour.MotionEffect.ToX = 0;
                behaviour.MotionEffect.ToY = (fontSize / fontScale) * (currentMultiplier - 1);
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant move up effects for whitespaces and modifies effect animation to simulate a move up animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <param name="currentMultiplier">multiplier for the distance that the effect must travel</param>
        /// <param name="fontSize">size of font to scale the effects movement distance</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatMoveUpWhitespaceEffects(int lineToKeep, List<PowerPoint.Effect> effectList, int currentMultiplier, float fontSize)
        {
            if (lineToKeep > 0)
            {
                for (int i = lineToKeep-1; i >= 0; --i)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }


            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.5f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                PowerPoint.AnimationBehavior behaviour = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.FromY = (fontSize / fontScale) * currentMultiplier;
                behaviour.MotionEffect.ToX = 0;
                behaviour.MotionEffect.ToY = (fontSize / fontScale) * (currentMultiplier - 1);
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant move down effects and modifies effect animation to simulate a move down animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <param name="currentMultiplier">multiplier for the distance that the effect must travel</param>
        /// <param name="fontSize">size of font to scale the effects movement distance</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatMoveDownEffects(int lineToKeep, List<PowerPoint.Effect> effectList, int currentMultiplier, float fontSize)
        {
            if (lineToKeep > 0)
            {
                for (int i = lineToKeep - 1; i >= 0; --i)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }

            for (int effect = 0; effect < effectList.Count; effect++)
            {
                effectList[effect].Timing.Duration = 0.5f;
                PowerPoint.AnimationBehavior behaviour = effectList[effect].Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.FromY = (fontSize / fontScale) * currentMultiplier;
                behaviour.MotionEffect.ToX = 0;
                behaviour.MotionEffect.ToY = (fontSize / fontScale) * (currentMultiplier + 1);
                if (effect == 0)
                {
                    effectList[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                }
                else
                {
                    effectList[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                }
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant move down effects with whitespaces and modifies effect animation to simulate a move down animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <param name="currentMultiplier">multiplier for the distance that the effect must travel</param>
        /// <param name="fontSize">size of font to scale the effects movement distance</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatMoveDownWhitespaceEffects(int lineToKeep, List<PowerPoint.Effect> effectList, int currentMultiplier, float fontSize)
        {
            if (lineToKeep > 0)
            {
                for (int i = lineToKeep - 1; i >= 0; --i)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }

            for (int effect = 0; effect < effectList.Count; effect++)
            {
                effectList[effect].Timing.Duration = 0.5f;
                PowerPoint.AnimationBehavior behaviour = effectList[effect].Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.FromY = (fontSize / fontScale) * currentMultiplier;
                behaviour.MotionEffect.ToX = 0;
                behaviour.MotionEffect.ToY = (fontSize / fontScale) * (currentMultiplier + 1);
                if (effect == 0)
                {
                    effectList[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                }
                else
                {
                    effectList[effect].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                }
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant insert effects and modifies effect animation to simulate an insert animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatInsertEffects(int lineToKeep, List<PowerPoint.Effect> effectList)
        {
            for (int i = effectList.Count - 1; i >= 0; --i)
            {
                // delete redundant insert effects from back.
                if (i != lineToKeep)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                effect.Timing.Duration = 0.7f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant colour change effects for lines to disappear and modifies effect animation to simulate a colour change animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatDisappearColourChangeEffects(int lineToKeep, List<PowerPoint.Effect> effectList)
        {
            for (int i = effectList.Count - 1; i >= 0; --i)
            {
                // delete redundant colour change effects from back.
                if (i != lineToKeep)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.5f;
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            return effectList;
        }

        /// <summary>
        /// Removes all redundant colour change effects for lines to appear and modifies effect animation to simulate a colour change animation.
        /// </summary>
        /// <param name="lineToKeep">effect line which is retained while all other effects are deleted</param>
        /// <param name="effectList">list of effects to be formatted</param>
        /// <returns>list of formatted effects</returns>
        private List<PowerPoint.Effect> FormatAppearColourChangeEffects(int lineToKeep, List<PowerPoint.Effect> effectList)
        {
            for (int i = effectList.Count - 1; i >= 0; --i)
            {
                // delete redundant colour change effects from back.
                if (i != lineToKeep)
                {
                    effectList[i].Delete();
                    effectList.RemoveAt(i);
                }
            }
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.1f;
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
            return effectList;
        }

        /// <summary>
        /// Apply formatting and timing to the "disappear" effects (i.e. repetitive code).
        /// </summary>
        /// <param name="disappearEffects">list of disappear effects</param>
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
        /// Rearranges the effects to simulate code change in one single block
        /// </summary>
        /// <param name="effectList">list of effects to be rearranged</param>
        /// <param name="beforeEffect">effect to be queued in front of the effect list</param>
        /// <param name="triggerType"></param>
        private static void RearrangeBlockDiffEffects(List<PowerPoint.Effect> effectList, PowerPoint.Effect beforeEffect, PowerPoint.MsoAnimTriggerType triggerType)
        {
            for (int i = 0; i < effectList.Count; i++)
            {
                if (i == 0)
                {
                    effectList[i].Timing.TriggerType = triggerType;
                    effectList[i].MoveAfter(beforeEffect);
                    continue;
                }
                effectList[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                effectList[i].MoveAfter(effectList[i - 1]);
            }
        }
    }
}
