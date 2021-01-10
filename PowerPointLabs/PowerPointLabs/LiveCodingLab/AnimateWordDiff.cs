using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DiffMatchPatch;
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
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.LiveCodingLab
{
    public partial class LiveCodingLabMain
    {
#pragma warning disable 0618
        internal const int AnimateWordDiff_MinNoOfShapesRequired = 1;
        internal const string AnimateWordDiff_FeatureName = "Animate Word Diff";
        internal const string AnimateWordDiff_ShapeSupport = "code box";
        internal static readonly string[] AnimateWordDiff_ErrorParameters =
        {
            AnimateWordDiff_FeatureName,
            AnimateWordDiff_MinNoOfShapesRequired.ToString(),
            AnimateWordDiff_ShapeSupport
        };
        public void AnimateWordDiff(List<CodeBoxPaneItem> codeListBox)
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

                if (diffCodeBoxBefore.CodeBox.IsDiff && diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    if (diffCodeBoxBefore.CodeBox.Text != diffCodeBoxAfter.CodeBox.Text)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                        return;
                    }

                }
                else if (!diffCodeBoxBefore.CodeBox.IsDiff && !diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    // Check that there exists a "before" code and an "after" code to be animated
                    if (diffCodeBoxBefore.CodeBox.Shape == null || diffCodeBoxAfter.CodeBox.Shape == null)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                        return;
                    }

                    if (diffCodeBoxBefore.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse ||
                        diffCodeBoxAfter.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                        return;
                    }

                    diffCodeBoxAfter.CodeBox.Shape.Left = diffCodeBoxBefore.CodeBox.Shape.Left;
                    diffCodeBoxAfter.CodeBox.Shape.Top = diffCodeBoxBefore.CodeBox.Shape.Top;
                    diffCodeBoxAfter.CodeBox.Shape.Width = diffCodeBoxBefore.CodeBox.Shape.Width;
                    diffCodeBoxAfter.CodeBox.Shape.Height = diffCodeBoxBefore.CodeBox.Shape.Height;
                }
                else
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                // Creates a new animation slide between the before and after code
                PowerPointSlide transitionSlide = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlide.Index + 1);
                transitionSlide.Name = LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");




                IEnumerable<Tuple<WordDiffType, Shape, Shape>> transitionText = CreateTransitionTextForWordDiff(transitionSlide, diffCodeBoxBefore, diffCodeBoxAfter);

                CreateAnimationForTransitionText(transitionSlide, transitionText);

                AddPowerPointLabsIndicator(transitionSlide);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateWordDiff");
                throw;
            }
        }

        private void CreateAnimationForTransitionText(PowerPointSlide transitionSlide, IEnumerable<Tuple<WordDiffType, Shape, Shape>> transitionText)
        {
            PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;
            Dictionary<float, List<Shape>> shapesByLine = GetShapesByLine(transitionSlide);
            HashSet<Shape> addShapes = new HashSet<Shape>();
            float emptyTextboxOffset = 7.272875f;

            foreach (Tuple<WordDiffType, Shape, Shape> shapePair in transitionText)
            {
                if (shapePair.Item1 == WordDiffType.AddEqual)
                {
                    addShapes.Add(shapePair.Item2);
                }
                else if (shapePair.Item1 == WordDiffType.DeleteAdd)
                {
                    addShapes.Add(shapePair.Item3);
                }
            }

            foreach (Tuple<WordDiffType, Shape, Shape> pairToAnimate in transitionText)
            {
                int currentIndex = sequence.Count;

                Shape beforeShape = pairToAnimate.Item2;
                Shape afterShape = pairToAnimate.Item3;

                switch (pairToAnimate.Item1)
                {
                    case WordDiffType.AddEqual:
                        if (beforeShape.Top.Equals(afterShape.Top))
                        {
                            List<Shape> shapesToShift = shapesByLine[afterShape.Top];
                            int index = shapesToShift.IndexOf(beforeShape);
                            bool shiftShapesRight = false;
                            bool shiftShapesLeft = false;
                            float offset = 0.0f;

                            for (int i = index + 1; i < shapesToShift.Count; i++)
                            {
                                Shape shape = shapesToShift[i];
                                currentIndex = sequence.Count;

                                if (!addShapes.Contains(shape) && shiftShapesRight)
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathRight,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftRightEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveRightEffects(shiftRightEffects, offset);
                                }
                                else if (!addShapes.Contains(shape) && shiftShapesLeft)
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathLeft,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftLeftEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveLeftEffects(shiftLeftEffects, offset);
                                }
                                else if (!addShapes.Contains(shape) && shape.Left + emptyTextboxOffset < (beforeShape.Left + beforeShape.Width - emptyTextboxOffset))
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathRight,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftRightEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveRightEffects(shiftRightEffects,
                                        ((beforeShape.Left + beforeShape.Width - emptyTextboxOffset) - (shape.Left + emptyTextboxOffset)) / 10);
                                    shiftShapesRight = true;
                                    offset = ((beforeShape.Left + beforeShape.Width - emptyTextboxOffset) - (shape.Left + emptyTextboxOffset)) / 10;
                                }
                                else if (!addShapes.Contains(shape) && shape.Left + emptyTextboxOffset > (beforeShape.Left + beforeShape.Width - emptyTextboxOffset))
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathLeft,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftLeftEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveLeftEffects(shiftLeftEffects,
                                        ((shape.Left + emptyTextboxOffset) - (beforeShape.Left + beforeShape.Width - emptyTextboxOffset)) / 10);
                                    offset = ((shape.Left + emptyTextboxOffset) - (beforeShape.Left + beforeShape.Width - emptyTextboxOffset)) / 10;
                                }
                            }
                        }

                        currentIndex = sequence.Count;

                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> insertAddEqualEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffAddEffects(insertAddEqualEffects);

                        currentIndex = sequence.Count;

                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        List<PowerPoint.Effect> colourChangeAddEqualEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffColourChangeEffects(colourChangeAddEqualEffects);

                        break;
                    case WordDiffType.DeleteEqual:
                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeDeleteEqualEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffColourChangeEffects(colourChangeDeleteEqualEffects);

                        currentIndex = sequence.Count;

                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        List<PowerPoint.Effect> deleteEqualEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffDeleteEffects(deleteEqualEffects);

                        if (!beforeShape.Equals(afterShape) && beforeShape.Top.Equals(afterShape.Top))
                        {
                            List<Shape> shapesToShift = shapesByLine[afterShape.Top];
                            int index = shapesToShift.IndexOf(beforeShape);
                            for (int i = index + 1; i < shapesToShift.Count; i++)
                            {
                                Shape shape = shapesToShift[i];
                                currentIndex = sequence.Count;

                                sequence.AddEffect(shape,
                                    PowerPoint.MsoAnimEffect.msoAnimEffectPathLeft,
                                    PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                List<PowerPoint.Effect> deleteEqualMoveLeftEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                FormatWordDiffMoveLeftEffects(deleteEqualMoveLeftEffects, beforeShape.TextFrame.TextRange.Length);
                            }
                        }

                        break;
                    case WordDiffType.DeleteAdd:
                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        List<PowerPoint.Effect> colourChangeDeleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffColourChangeEffects(colourChangeDeleteEffects);

                        currentIndex = sequence.Count;

                        sequence.AddEffect(beforeShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        List<PowerPoint.Effect> deleteEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffDeleteEffects(deleteEffects);

                        if (beforeShape.Top.Equals(afterShape.Top))
                        {
                            List<Shape> shapesToShift = shapesByLine[afterShape.Top];
                            int index = shapesToShift.IndexOf(afterShape);
                            bool shiftShapesRight = false;
                            bool shiftShapesLeft = false;
                            float offset = 0.0f;

                            for (int i = index + 1; i < shapesToShift.Count; i++)
                            {
                                Shape shape = shapesToShift[i];

                                currentIndex = sequence.Count;

                                if (!addShapes.Contains(shape) && shiftShapesRight)
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathRight,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftRightEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveRightEffects(shiftRightEffects, offset);

                                }
                                else if (!addShapes.Contains(shape) && shiftShapesLeft)
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathLeft,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftLeftEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveLeftEffects(shiftLeftEffects, offset);
                                }
                                else if (!addShapes.Contains(shape) && shape.Left + emptyTextboxOffset < (afterShape.Left + afterShape.Width - emptyTextboxOffset))
                                {
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathRight,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftRightEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveRightEffects(shiftRightEffects, 
                                        ((afterShape.Left + afterShape.Width - emptyTextboxOffset) - (shape.Left + emptyTextboxOffset)) / 10);
                                    shiftShapesRight = true;
                                    offset = ((afterShape.Left + afterShape.Width - emptyTextboxOffset) - (shape.Left + emptyTextboxOffset)) / 10;
                                }
                                else if (!addShapes.Contains(shape) && shape.Left + emptyTextboxOffset > (afterShape.Left + afterShape.Width - emptyTextboxOffset))
                                {
                                    MessageBox.Show((shape.Left + emptyTextboxOffset).ToString(), (afterShape.Left + afterShape.Width - emptyTextboxOffset).ToString());
                                    sequence.AddEffect(shape,
                                        PowerPoint.MsoAnimEffect.msoAnimEffectPathLeft,
                                        PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel,
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    List<PowerPoint.Effect> shiftLeftEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                                    FormatWordDiffMoveLeftEffects(shiftLeftEffects,
                                        ((shape.Left + emptyTextboxOffset) - (afterShape.Left + afterShape.Width - emptyTextboxOffset)) / 10);
                                    shiftShapesLeft = true;
                                    offset = ((shape.Left + emptyTextboxOffset) - (afterShape.Left + afterShape.Width - emptyTextboxOffset)) / 10;
                                }
                            }
                        }

                        currentIndex = sequence.Count;

                        sequence.AddEffect(afterShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByAllLevels,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        List<PowerPoint.Effect> addEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffAddEffects(addEffects);

                        currentIndex = sequence.Count;

                        sequence.AddEffect(afterShape,
                            PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
                            PowerPoint.MsoAnimateByLevel.msoAnimateTextByAllLevels,
                            PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        List<PowerPoint.Effect> colourChangeAddEffects = AsList(sequence, currentIndex + 1, sequence.Count + 1);
                        FormatWordDiffColourChangeEffects(colourChangeAddEffects);

                        break;
                    default:
                        break;
                }
            }
        }

        private List<Tuple<WordDiffType, Shape, Shape>> CreateTransitionTextForWordDiff(PowerPointSlide transitionSlide, CodeBoxPaneItem diffCodeBoxBefore, CodeBoxPaneItem diffCodeBoxAfter)
        {
            PowerPoint.TextRange codeTextBeforeEdit = diffCodeBoxBefore.CodeBox.Shape.TextFrame.TextRange;
            PowerPoint.TextRange codeTextAfterEdit = diffCodeBoxAfter.CodeBox.Shape.TextFrame.TextRange;

            var differ = DiffMatchPatchModule.Default;
            var diffs = differ.DiffMain(codeTextBeforeEdit.Text, codeTextAfterEdit.Text);
            differ.DiffCleanupSemantic(diffs);
            float emptyTextboxOffset = 7.272875f;
            float topPointerLineOffset = 20 * (diffCodeBoxBefore.CodeBox.Shape.TextFrame.TextRange.Font.Size / 18);
            float originalLeftPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float leftBeforePointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float leftEqualPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float topBeforePointer = diffCodeBoxBefore.CodeBox.Shape.Top;
            float leftAfterPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float topAfterPointer = diffCodeBoxBefore.CodeBox.Shape.Top;
            List<Tuple<Operation, Shape>> transitionText = new List<Tuple<Operation, Shape>>();
            List<Tuple<WordDiffType, Shape, Shape>> transitionTextToAnimate = new List<Tuple<WordDiffType, Shape, Shape>>();

            for (int j = 0; j < diffs.Count; j++)
            {
                string text = diffs[j].Text;
                string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                float leftPointer;
                float topPointer;

                if (diffs[j].Operation.IsInsert)
                {
                    leftPointer = leftAfterPointer;
                    topPointer = topAfterPointer;
                }
                else
                {
                    leftPointer = leftBeforePointer;
                    topPointer = topBeforePointer;
                }

                Shape textbox = transitionSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    leftPointer, topPointer, 0, 0);
                
                textbox.TextFrame.TextRange.Text = lines[0];
                textbox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                textbox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                textbox.TextFrame.TextRange.Font.Size = codeTextBeforeEdit.Font.Size;
                textbox.TextFrame.TextRange.Font.Name = codeTextBeforeEdit.Font.Name;
                textbox.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
                textbox.TextEffect.Alignment = Office.MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
                transitionText.Add(Tuple.Create(diffs[j].Operation, textbox));

                leftPointer += textbox.Width - (2 * emptyTextboxOffset);
                if (!diffs[j].Operation.IsDelete)
                {
                    leftEqualPointer += textbox.Width - (2 * emptyTextboxOffset);
                }
                if (lines[0].Contains("\n"))
                {
                    leftPointer = originalLeftPointer;
                    leftEqualPointer = originalLeftPointer;
                    topPointer += topPointerLineOffset;
                }

                if (lines.Length > 1)
                {
                    for (int i = 1; i < lines.Length; i++)
                    {
                        leftPointer = originalLeftPointer;
                        leftEqualPointer = originalLeftPointer;
                        topPointer += topPointerLineOffset;

                        textbox = transitionSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            leftPointer, topPointer, 0, 0);

                        textbox.TextFrame.TextRange.Text = lines[i];
                        textbox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        textbox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                        textbox.TextFrame.TextRange.Font.Size = codeTextBeforeEdit.Font.Size;
                        textbox.TextFrame.TextRange.Font.Name = codeTextBeforeEdit.Font.Name;
                        textbox.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
                        textbox.TextEffect.Alignment = Office.MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
                        transitionText.Add(Tuple.Create(diffs[j].Operation, textbox));
                    }
                    leftPointer += textbox.Width - (2 * emptyTextboxOffset);
                    if (!diffs[j].Operation.IsDelete)
                    {
                        leftEqualPointer += textbox.Width - (2 * emptyTextboxOffset);
                    }
                    if (lines[lines.Length - 1].Contains("\n"))
                    {
                        leftPointer = originalLeftPointer;
                        leftEqualPointer = originalLeftPointer;
                        topPointer += topPointerLineOffset;
                    }
                }
                if (diffs[j].Operation.IsDelete)
                {
                    leftBeforePointer = leftPointer;
                    topBeforePointer = topPointer;
                }
                else if (diffs[j].Operation.IsInsert)
                {
                    if (topAfterPointer != topPointer)
                    {
                        topBeforePointer = topPointer;
                    }
                    leftAfterPointer = leftPointer;
                    topAfterPointer = topPointer;
                }
                else
                {
                    leftAfterPointer = leftEqualPointer;
                    topAfterPointer += topPointer - topBeforePointer;
                    leftBeforePointer = leftPointer;
                    topBeforePointer = topPointer;
                }
            }

            for (int i = transitionText.Count - 1; i >= 0; i--)
            {
                if (transitionText[i].Item2.TextFrame.TextRange.Length == 0)
                {
                    transitionText.RemoveAt(i);
                }
            }

            for (int i = 1; i < transitionText.Count; i++)
            {
                WordDiffType transitionTextType;
                if (transitionText[i-1].Item1 == Operation.Delete && transitionText[i].Item1 == Operation.Insert)
                {
                    transitionTextType = WordDiffType.DeleteAdd;
                }
                else if (transitionText[i-1].Item1 == Operation.Delete)
                {
                    transitionTextType = WordDiffType.DeleteEqual;
                }
                else if (transitionText[i-1].Item1 == Operation.Insert)
                {
                    transitionTextType = WordDiffType.AddEqual;
                }
                else
                {
                    transitionTextType = WordDiffType.Equal;
                }
                if (transitionTextType != WordDiffType.Equal)
                {
                    transitionTextToAnimate.Add(Tuple.Create(transitionTextType, transitionText[i-1].Item2, transitionText[i].Item2));
                }
                if (transitionTextType == WordDiffType.DeleteAdd)
                {
                    i++;
                }
            }
            if (transitionText[transitionText.Count-1].Item1 == Operation.Delete)
            {
                transitionTextToAnimate.Add(Tuple.Create(WordDiffType.DeleteEqual, transitionText[transitionText.Count - 1].Item2, transitionText[transitionText.Count - 1].Item2));
            }
            else if (transitionText[transitionText.Count - 1].Item1 == Operation.Insert)
            {
                transitionTextToAnimate.Add(Tuple.Create(WordDiffType.AddEqual, transitionText[transitionText.Count - 1].Item2, transitionText[transitionText.Count - 1].Item2));
            }
            return transitionTextToAnimate;
        }

        private void FormatWordDiffDeleteEffects(List<PowerPoint.Effect> effectList)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionRight;
                effect.Exit = Office.MsoTriState.msoTrue;
                effect.Timing.Duration = 0.7f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        private void FormatWordDiffAddEffects(List<PowerPoint.Effect> effectList)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                effect.Timing.Duration = 0.7f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        private void FormatWordDiffColourChangeEffects(List<PowerPoint.Effect> effectList)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.1f;
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
            }
        }

        private void FormatWordDiffMoveLeftEffects(List<PowerPoint.Effect> effectList, float offset)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.5f;
                PowerPoint.AnimationBehavior behaviour = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.ToX = -offset;
                behaviour.MotionEffect.FromY = 0;
                behaviour.MotionEffect.ToY = 0;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        private void FormatWordDiffMoveRightEffects(List<PowerPoint.Effect> effectList, float offset)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.3f;
                PowerPoint.AnimationBehavior behaviour = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behaviour.MotionEffect.FromX = 0;
                behaviour.MotionEffect.ToX = offset;
                behaviour.MotionEffect.FromY = 0;
                behaviour.MotionEffect.ToY = 0;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        private Dictionary<float, List<Shape>> GetShapesByLine(PowerPointSlide slide)
        {
            Dictionary<float, List<Shape>> shapesByLine = new Dictionary<float, List<Shape>>();

            foreach (Shape shape in slide.Shapes)
            {
                float shapeTopPosition = shape.Top;
                
                if (shapesByLine.ContainsKey(shapeTopPosition))
                {
                    shapesByLine[shapeTopPosition].Add(shape);
                }
                else
                {
                    shapesByLine.Add(shapeTopPosition, new List<Shape>() { shape });
                }
            }

            return shapesByLine;
        }
    }
}
