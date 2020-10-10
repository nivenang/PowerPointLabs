using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
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
        public void InsertDiff(List<FileDiff> diffList, LiveCodingPaneWPF parent, string diffGroup)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }
                
                if (diffList.Count < 1) 
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }
                
                FileDiff diff = diffList[0];

                PowerPointSlide diffSlideBefore = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlide.Index+1);
                // INSERT LOGIC FOR INSERTING SHAPE INTO SLIDE
                // CREATE CODEBOXPANEITEM FROM DIFF
                List<ChunkDiff> diffChunks = diff.Chunks.ToList();
                
                string codeTextBefore = "";
                string codeTextAfter = "";

                List<int> markedForDisappear = new List<int>();
                List<int> markedForAppear = new List<int>();
                int beforeCounter = 0;
                int afterCounter = 0;
                foreach (ChunkDiff chunk in diffChunks)
                {
                    List<LineDiff> diffLines = chunk.Changes.ToList();
                    foreach (LineDiff line in diffLines)
                    {
                        if (line.Add)
                        {
                            codeTextAfter += AppendLineEnd(line.Content.Trim().TrimStart('+'));
                            markedForAppear.Add(afterCounter);
                            afterCounter++;
                        }
                        else if (line.Delete)
                        {
                            codeTextBefore += AppendLineEnd(line.Content.Trim().TrimStart('-'));
                            markedForDisappear.Add(beforeCounter);
                            beforeCounter++;
                        }
                        else
                        {
                            codeTextBefore += AppendLineEnd(line.Content.Trim());
                            codeTextAfter += AppendLineEnd(line.Content.Trim());
                            beforeCounter++;
                            afterCounter++;
                        }
                    }
                    codeTextBefore += "...\r\n";
                    codeTextAfter += "...\r\n";
                    beforeCounter++;
                    afterCounter++;
                }

                CodeBoxPaneItem codeBoxPaneItemBefore = new CodeBoxPaneItem(parent);
                CodeBoxPaneItem codeBoxPaneItemAfter = new CodeBoxPaneItem(parent);

                // UPDATE CODEBOX TO HAVE NEW TEXT
                codeBoxPaneItemBefore.CodeBox.Text = codeTextBefore;
                codeBoxPaneItemBefore.Group = diffGroup;
                codeBoxPaneItemAfter.CodeBox.Text = codeTextAfter;
                codeBoxPaneItemAfter.Group = diffGroup;

                parent.AddCodeBox(codeBoxPaneItemBefore);
                parent.AddCodeBox(codeBoxPaneItemAfter);

                CodeBox diffCodeBoxBefore = ShapeUtility.InsertCodeBoxToSlide(diffSlideBefore, codeBoxPaneItemBefore.CodeBox);

                PowerPointSlide diffSlideAfter = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, "", currentSlide.Index + 2);
                // INSERT LOGIC FOR INSERTING SHAPE INTO SLIDE
                CodeBox diffCodeBoxAfter = ShapeUtility.InsertCodeBoxToSlide(diffSlideAfter, codeBoxPaneItemAfter.CodeBox);

                codeBoxPaneItemBefore.CodeBox = diffCodeBoxBefore;
                codeBoxPaneItemAfter.CodeBox = diffCodeBoxAfter;
                // CREATE ANIMATION FOR DIFFERENCES BETWEEN TWO SLIDES

                
                // Creates a new animation slide between the before and after code
                PowerPointSlide transitionSlide = diffSlideBefore.Duplicate();
                transitionSlide.Name = "PPTLabsInsertDiffTransitionSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                AddPowerPointLabsIndicator(transitionSlide);

                // Initialise an animation sequence object
                PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;

                PowerPoint.Shape codeShapeBeforeEdit = transitionSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex)[0];
                PowerPoint.Shape codeShapeAfterEdit = transitionSlide.CopyShapeToSlide(diffCodeBoxAfter.Shape);
                codeShapeBeforeEdit = ConvertTextToParagraphs(codeShapeBeforeEdit);
                codeShapeAfterEdit = ConvertTextToParagraphs(codeShapeAfterEdit);
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
                while (beforeCount < codeTextBeforeEdit.Paragraphs().Count && afterCount < codeTextAfterEdit.Paragraphs().Count)
                {
                    if (codeTextBeforeEdit.Paragraphs(beforeCount+1).TrimText().Text == "")
                    {
                        if (markedForDisappear.Contains(beforeCount))
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
                        continue;
                    }
                    
                    if (codeTextAfterEdit.Paragraphs(afterCount+1).TrimText().Text == "")
                    {
                        if (markedForAppear.Contains(afterCount))
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
                        continue;
                    }
                    
                    if (markedForDisappear.Contains(beforeCount))
                    {
                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> deleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        deleteEffects = FormatDeleteEffects(beforeLineToEffectLine[beforeCount], deleteEffects);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectPathUp,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        List<PowerPoint.Effect> moveUpEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        moveUpEffects = FormatMoveUpEffects(beforeLineToEffectLine[beforeCount], moveUpEffects, currentMultiplier, fontSize);
                        
                        currentMultiplier--;
                        beforeCount++;
                    }
                    else if (markedForAppear.Contains(afterCount))
                    {
                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeBeforeEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        List<PowerPoint.Effect> moveDownEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        moveDownEffects = FormatMoveDownEffects(beforeLineToEffectLine[beforeCount], moveDownEffects, currentMultiplier, fontSize);

                        currentIndex = sequence.Count;
                        sequence.AddEffect(codeShapeAfterEdit,
                            PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> insertEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);

                        insertEffects = FormatInsertEffects(afterLineToEffectLine[afterCount], insertEffects);

                        currentMultiplier++;
                        afterCount++;
                    }
                    else
                    {
                        beforeCount++;
                        afterCount++;
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
                Logger.LogException(e, "InsertDiff");
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

        private string AppendLineEnd(string line)
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
