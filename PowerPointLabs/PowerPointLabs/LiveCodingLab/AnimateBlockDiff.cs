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
        internal const int AnimateBlockDiff_MinNoOfShapesRequired = 1;
        internal const string AnimateBlockDiff_FeatureName = "Animate Block Diff";
        internal const string AnimateBlockDiff_ShapeSupport = "code box";
        internal static readonly string[] AnimateBlockDiff_ErrorParameters =
        {
            AnimateBlockDiff_FeatureName,
            AnimateBlockDiff_MinNoOfShapesRequired.ToString(),
            AnimateBlockDiff_ShapeSupport
        };
        public void AnimateBlockDiff(List<CodeBoxPaneItem> codeListBox)
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

                AnimateDiffByBlock(codeListBox, diffList[0]);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateBlockDiff");
                throw;
            }
        }

        private void AnimateDiffByBlock (List<CodeBoxPaneItem> codeListBox, FileDiff diff)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

                CodeBoxPaneItem diffCodeBoxBefore = codeListBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = codeListBox[1];

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
                    if (codeTextBeforeEdit.Paragraphs(i + 1).TrimText().Text == "")
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
                List<PowerPoint.Effect> disappearHighlightEffects = new List<PowerPoint.Effect>();
                List<PowerPoint.Effect> appearHighlightEffects = new List<PowerPoint.Effect>();
                List<PowerPoint.Effect> intermediateEffects = new List<PowerPoint.Effect>();
                
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

                                intermediateEffects.AddRange(moveUpEffects);

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

                        disappearHighlightEffects.AddRange(colourChangeEffectsBefore);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> deleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        deleteEffects = FormatDeleteEffects(beforeLineToEffectLine[beforeCount], deleteEffects);

                        intermediateEffects.AddRange(deleteEffects);

                        if (lineCount + 1 >= fullDiff.Count || (lineCount + 1 < fullDiff.Count && fullDiff[lineCount + 1] != DiffType.Add))
                        {
                            currentIndex = sequence.Count;
                            sequence.AddEffect(codeShapeBeforeEdit,
                                PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                                PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                            moveUpEffects = FormatMoveUpEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);

                            intermediateEffects.AddRange(moveUpEffects);

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

                                intermediateEffects.AddRange(moveDownEffects);

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

                            intermediateEffects.AddRange(moveDownEffects);

                            currentMultiplier++;
                        }

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> insertEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        insertEffects = FormatInsertEffects(afterLineToEffectLine[afterCount], insertEffects);

                        intermediateEffects.AddRange(insertEffects);

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
                    else
                    {
                        beforeCount++;
                        afterCount++;
                        lineCount++;
                    }
                }

                for (int i = 0; i < disappearHighlightEffects.Count; i++)
                {
                    if (i == 0)
                    {
                        disappearHighlightEffects[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                        disappearHighlightEffects[i].MoveAfter(disappearEffects[disappearEffects.Count - 1]);
                        continue;
                    }
                    disappearHighlightEffects[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    disappearHighlightEffects[i].MoveAfter(disappearHighlightEffects[i - 1]);
                }

                for (int i = 0; i < intermediateEffects.Count; i++)
                {
                    if (i == 0)
                    {
                        intermediateEffects[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                        intermediateEffects[i].MoveAfter(disappearHighlightEffects[disappearHighlightEffects.Count - 1]);
                        continue;
                    }
                    intermediateEffects[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    intermediateEffects[i].MoveAfter(intermediateEffects[i - 1]);
                }

                for (int i = 0; i < appearHighlightEffects.Count; i++)
                {
                    appearHighlightEffects[i].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    if (i == 0)
                    {
                        appearHighlightEffects[i].MoveAfter(intermediateEffects[intermediateEffects.Count - 1]);
                        continue;
                    }
                    appearHighlightEffects[i].MoveAfter(appearHighlightEffects[i - 1]);
                }

                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }
                PowerPointPresentation.Current.AddAckSlide();
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateDiffByBlock");
                throw;
            }
        }
    }
}
