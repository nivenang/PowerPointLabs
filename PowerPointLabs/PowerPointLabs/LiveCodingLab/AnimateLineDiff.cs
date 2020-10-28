﻿using System;
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
        internal const int AnimateLineDiff_MinNoOfShapesRequired = 1;
        internal const string AnimateLineDiff_FeatureName = "Animate Line Diff";
        internal const string AnimateLineDiff_ShapeSupport = "code box";
        internal static readonly string[] AnimateLineDiff_ErrorParameters =
        {
            AnimateLineDiff_FeatureName,
            AnimateLineDiff_MinNoOfShapesRequired.ToString(),
            AnimateLineDiff_ShapeSupport
        };
        public void AnimateLineDiff(List<CodeBoxPaneItem> codeListBox)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }

                if (codeListBox.Count != 2)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                CodeBoxPaneItem diffCodeBoxBefore = codeListBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = codeListBox[1];

                if (!diffCodeBoxBefore.CodeBox.IsDiff || !diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                if (diffCodeBoxBefore.CodeBox.Text != diffCodeBoxAfter.CodeBox.Text)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                List<FileDiff> diffList = CodeBoxFileService.ParseDiff(diffCodeBoxBefore.CodeBox.Text);

                if (diffList.Count < 1)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                AnimateDiffByLine(codeListBox, diffList[0]);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateLineDiff");
                throw;
            }
        }
        private void AnimateDiffByLine (List<CodeBoxPaneItem> listCodeBox, FileDiff diff)
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
                transitionSlide.Name = LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                AddPowerPointLabsIndicator(transitionSlide);

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;

                PowerPoint.Shape codeShapeBeforeEdit = transitionSlide.CopyShapeToSlide(diffCodeBoxBefore.CodeBox.Shape);
                PowerPoint.Shape codeShapeAfterEdit = transitionSlide.CopyShapeToSlide(diffCodeBoxAfter.CodeBox.Shape);

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
                while (beforeCount < codeTextBeforeEdit.Paragraphs().Count && afterCount < codeTextAfterEdit.Paragraphs().Count)
                {
                    if (fullDiff[lineCount] == DiffType.Delete)
                    {
                        if (codeTextBeforeEdit.Paragraphs(beforeCount + 1).TrimText().Text == "")
                        {
                            if (lineCount + 1 >= fullDiff.Count || (lineCount + 1 < fullDiff.Count && fullDiff[lineCount + 1] != DiffType.Add))
                            {
                                currentIndex = sequence.Count;
                                sequence.AddEffect(codeShapeBeforeEdit,
                                    PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                                moveUpEffects = FormatMoveUpWhitespaceEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);

                                currentMultiplier--;
                            }
                            beforeCount++;
                            lineCount++;
                            continue;
                        }
                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeEffectsBefore = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        colourChangeEffectsBefore = FormatDisappearColourChangeEffects(beforeLineToEffectLine[beforeCount], colourChangeEffectsBefore);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> deleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        deleteEffects = FormatDeleteEffects(beforeLineToEffectLine[beforeCount], deleteEffects);

                        if (lineCount + 1 >= fullDiff.Count || (lineCount + 1 < fullDiff.Count && fullDiff[lineCount + 1] != DiffType.Add))
                        {
                            currentIndex = sequence.Count;
                            sequence.AddEffect(codeShapeBeforeEdit,
                                PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                                PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                            moveUpEffects = FormatMoveUpEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);
                            
                            currentMultiplier--;
                        }

                        beforeCount++;
                        lineCount++;
                    }
                    else if (fullDiff[lineCount] == DiffType.Add)
                    {
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

                                currentMultiplier++;
                            }
                            afterCount++;
                            lineCount++;
                            continue;
                        }

                        if (lineCount == 0 || (lineCount - 1 >= 0 && fullDiff[lineCount - 1] != DiffType.Delete))
                        {
                            currentIndex = sequence.Count;
                            sequence.AddEffect(codeShapeBeforeEdit,
                                PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                                PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            List<PowerPoint.Effect> moveDownEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                            moveDownEffects = FormatMoveDownEffects(beforeLineToEffectLine[beforeCount], moveDownEffects, currentMultiplier, fontSize);
                            currentMultiplier++;
                        }

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> insertEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        insertEffects = FormatInsertEffects(afterLineToEffectLine[afterCount], insertEffects);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeEffectsAfter = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        colourChangeEffectsAfter = FormatAppearColourChangeEffects(afterLineToEffectLine[afterCount], colourChangeEffectsAfter);

                        afterCount++;
                        lineCount++;
                    }
                    else
                    {
                        beforeCount++;
                        afterCount++;
                        lineCount++;
                    }
                }
                
                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateDiffByLine");
                throw;
            }
        }

        /// <summary>
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
                effect.Exit = Office.MsoTriState.msoTrue;
                effect.Timing.Duration = 0.5f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            return effectList;
        }
        /// <summary>
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        /// Deletes all redundant effects from the sequence.
        /// </summary>
        private List<PowerPoint.Effect> FormatInsertEffects(int lineToKeep, List<PowerPoint.Effect> effectList)
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
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
            return effectList;
        }
        /// <summary>
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        /// Deletes all redundant effects from the sequence.
        /// </summary>
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
        private static void FormatDisappearEffects(List<PowerPoint.Effect> disappearEffects)
        {
            foreach (PowerPoint.Effect effect in disappearEffects)
            {
                effect.Exit = Office.MsoTriState.msoTrue;
                effect.Timing.Duration = 0;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
        }
    }
}
