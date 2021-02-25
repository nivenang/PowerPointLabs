using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
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
                PowerPointSlide nextSlide = currentPresentation.GetSlide(currentSlide.Index + 1);

                // Check that there is a slide selected by the user
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }

                // Check that there exists a "before" and "after" code
                if (codeListBox.Count != 2)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateWordDiffDialogTitle);
                    return;
                }

                CodeBoxPaneItem diffCodeBoxBefore = codeListBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = codeListBox[1];

                // Case 1: Animating differences across a Diff File
                if (diffCodeBoxBefore.CodeBox.IsDiff && diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    if (diffCodeBoxBefore.CodeBox.Text != diffCodeBoxAfter.CodeBox.Text)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffWrongCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateWordDiffDialogTitle);
                        return;
                    }

                }
                // Case 2: Animating differences across two user-input code snippets by building a diff file
                else if (!diffCodeBoxBefore.CodeBox.IsDiff && !diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    // Check that there exists a "before" code and an "after" code to be animated
                    if (diffCodeBoxBefore.CodeBox.Shape == null || diffCodeBoxAfter.CodeBox.Shape == null)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateWordDiffDialogTitle);
                        return;
                    }

                    if (diffCodeBoxBefore.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse ||
                        diffCodeBoxAfter.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateWordDiffDialogTitle);
                        return;
                    }

                    diffCodeBoxAfter.CodeBox.Shape.Left = diffCodeBoxBefore.CodeBox.Shape.Left;
                    diffCodeBoxAfter.CodeBox.Shape.Top = diffCodeBoxBefore.CodeBox.Shape.Top;
                    diffCodeBoxAfter.CodeBox.Shape.Width = diffCodeBoxBefore.CodeBox.Shape.Width;
                    diffCodeBoxAfter.CodeBox.Shape.Height = diffCodeBoxBefore.CodeBox.Shape.Height;
                }
                // Default: Inform user that code snippets to be animated do not match up
                else
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateWordDiffDialogTitle);
                    return;
                }

                // Creates a new animation slide between the before and after code
                // PowerPointSlide transitionSlide = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlide.Index + 1);
                PowerPointAutoAnimateSlide transitionSlide = AddTransitionAnimations(currentSlide, nextSlide);

                transitionSlide.Name = LiveCodingLabText.AnimateWordDiffIdentifier + LiveCodingLabText.TransitionSlideIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                // Create the transition text in the transition slide for Animating Word Diff
                IEnumerable<Tuple<WordDiffType, Shape, Shape>> transitionText = CreateTransitionTextForWordDiff(transitionSlide, diffCodeBoxBefore, diffCodeBoxAfter);

                // Animates the differences between the "before" and "after" code in the transition slide
                CreateAnimationForTransitionTextWordDiff(transitionSlide, transitionText);

                if (!transitionSlide.HasShapeWithRule(new Regex(@"PPTIndicator.*")))
                {
                    AddPowerPointLabsIndicator(transitionSlide);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateWordDiff");
                throw;
            }
        }

        /// <summary>
        /// Animates the differences between the "before" and "after" code snippets in the transition slide
        /// Precondition: shapes in the transition slide must exist
        /// </summary>
        /// <param name="transitionSlide">transition slide to animate the code snippets</param>
        /// <param name="transitionText">list containing tuples of shapes to create animations between</param>
        private void CreateAnimationForTransitionTextWordDiff(PowerPointSlide transitionSlide, IEnumerable<Tuple<WordDiffType, Shape, Shape>> transitionText)
        {
            PowerPoint.Sequence sequence = transitionSlide.TimeLine.MainSequence;
            Dictionary<float, List<Shape>> shapesByLine = GetShapesByLine(transitionSlide);
            HashSet<Shape> addShapes = new HashSet<Shape>();
            float emptyTextboxOffset = 7.272875f;

            // Initialise a hash set of all addition text boxes
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

            // Create animations for each textbox
            foreach (Tuple<WordDiffType, Shape, Shape> pairToAnimate in transitionText)
            {
                int currentIndex = sequence.Count;

                Shape beforeShape = pairToAnimate.Item2;
                Shape afterShape = pairToAnimate.Item3;

                switch (pairToAnimate.Item1)
                {
                    // Case 1: First textbox contains code to be added
                    case WordDiffType.AddEqual:

                        // Create movement effects if the following shape is on the same line
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

                        // Create appear effects for the addition code
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
                    // Case 2: First textbox contains deletion code, second textbox has deletion or equal code
                    case WordDiffType.DeleteEqual:
                        
                        // Create deletion animation effects for the deletion code
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

                        // Create the move left (closing up the gap from deleted line) effects only if following code is on the same line
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
                    // Case 3: First textbox contains deletion code, second textbox has addition code
                    case WordDiffType.DeleteAdd:

                        // Create deletion animation effects for the deletion code
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

                        // Create the movement effects to either close the gap from the deletion or to create space for the addition line
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

                        // Create addition animation effects for the new addition line
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
                    // Default: No animation created
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Creates the "before" and "after" code in the transition slide with each difference having a text box for animation
        /// </summary>
        /// <param name="transitionSlide">transition slide to create the code snippets in</param>
        /// <param name="diffCodeBoxBefore">code box containing the "before" code snippet</param>
        /// <param name="diffCodeBoxAfter">code box containing the "after" code snippet</param>
        /// <returns>list of shapes to be animated in order</returns>
        private List<Tuple<WordDiffType, Shape, Shape>> CreateTransitionTextForWordDiff(PowerPointSlide transitionSlide, CodeBoxPaneItem diffCodeBoxBefore, CodeBoxPaneItem diffCodeBoxAfter)
        {
            float emptyTextboxOffset = 7.272875f;
            float topPointerLineOffset = 20 * (diffCodeBoxBefore.CodeBox.Shape.TextFrame.TextRange.Font.Size / 18);
            float originalLeftPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float leftBeforePointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float leftEqualPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float topBeforePointer = diffCodeBoxBefore.CodeBox.Shape.Top;
            float leftAfterPointer = diffCodeBoxBefore.CodeBox.Shape.Left;
            float topAfterPointer = diffCodeBoxBefore.CodeBox.Shape.Top;
            int charCountBefore = 1;
            List<Tuple<ChangeType, Shape>> transitionText = new List<Tuple<ChangeType, Shape>>();
            List<Tuple<WordDiffType, Shape, Shape>> transitionTextToAnimate = new List<Tuple<WordDiffType, Shape, Shape>>();
            
            // Use Diff library to create differences in words across the "before" and "after" code
            PowerPoint.TextRange codeTextBeforeEdit = diffCodeBoxBefore.CodeBox.Shape.TextFrame.TextRange;
            PowerPoint.TextRange codeTextAfterEdit = diffCodeBoxAfter.CodeBox.Shape.TextFrame.TextRange;

            DiffPaneModel model = new DiffPaneModel();
            DiffPlex.Differ differ = new DiffPlex.Differ();

            DiffPlex.Model.DiffResult diffResult = differ.CreateWordDiffs(codeTextBeforeEdit.Text, codeTextAfterEdit.Text, true, true, new char[] { ',', ' ', '\'', '\n', '\r'});
            BuildDiffPieces(diffResult, model.Lines);
            var diffs = model.Lines;
            // Create individual textboxes for each diff object
            for (int j = 0; j < diffs.Count; j++)
            {
                // Split each diff based on newlines
                string text = diffs[j].Text;
                string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                // Maintain a left and top pointer for both "before" and "after code
                // to get the last known position of the code
                float leftPointer;
                float topPointer;

                // Use the "after" code pointer if the line to be created is an addition line
                if (diffs[j].Type == ChangeType.Inserted)
                {
                    leftPointer = leftAfterPointer;
                    topPointer = topAfterPointer;
                }
                // else, use the "before" code pointer if the line to be created is a deletion line
                else
                {
                    leftPointer = leftBeforePointer;
                    topPointer = topBeforePointer;
                }

                // Create a textbox for the first part of the code line (driver for code lines with > 1 line)
                Shape textbox = transitionSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    leftPointer, topPointer, 0, 0);

                textbox.Name = LiveCodingLabText.TransitionTextIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                textbox.TextFrame.TextRange.Text = lines[0];
                textbox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                textbox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                textbox.TextFrame.TextRange.Font.Size = codeTextBeforeEdit.Font.Size;
                textbox.TextFrame.TextRange.Font.Name = codeTextBeforeEdit.Font.Name;
                textbox.TextEffect.Alignment = Office.MsoTextEffectAlignment.msoTextEffectAlignmentLeft;

                // Add the created textbox reference to a list for further processing
                transitionText.Add(Tuple.Create(diffs[j].Type, textbox));

                // Syntax Highlighting for the created textbox
                for (int charIndex = 1; charIndex <= lines[0].Length; charIndex++)
                {
                    if (diffs[j].Type == ChangeType.Deleted || diffs[j].Type == ChangeType.Unchanged)
                    {
                        textbox.TextFrame.TextRange.Characters(charIndex, 1).Font.Color.RGB = codeTextBeforeEdit.Characters(charCountBefore, 1).Font.Color.RGB;
                        charCountBefore++;
                    }
                }

                // Adjust left pointer to the end of the newly created textbox for next line 
                leftPointer += textbox.Width - (2 * emptyTextboxOffset);

                // Increment the pointer for code lines containing only Equal lines if line is equal
                if (diffs[j].Type != ChangeType.Deleted)
                {
                    leftEqualPointer += textbox.Width - (2 * emptyTextboxOffset);
                }

                // Set pointers to next line if the code line ends with a newline
                if (lines[0].Contains("\n"))
                {
                    leftPointer = originalLeftPointer;
                    leftEqualPointer = originalLeftPointer;
                    topPointer += topPointerLineOffset;
                }

                // Increment Syntax Highlighting pointer if line ends with newline
                if (diffs[j].Type != ChangeType.Inserted && charCountBefore <= codeTextBeforeEdit.Characters().Length && Char.IsControl(codeTextBeforeEdit.Characters(charCountBefore, 1).Text, 0))
                {
                    charCountBefore++;
                }

                // Repeatedly create textboxes for the remaining part of the code line (for diff code lines > 1 line)
                if (lines.Length > 1)
                {
                    for (int i = 1; i < lines.Length; i++)
                    {
                        leftPointer = originalLeftPointer;
                        leftEqualPointer = originalLeftPointer;
                        topPointer += topPointerLineOffset;

                        // Create text box for the code line
                        textbox = transitionSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            leftPointer, topPointer, 0, 0);

                        textbox.Name = LiveCodingLabText.TransitionTextIdentifier + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                        textbox.TextFrame.TextRange.Text = lines[i];
                        textbox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        textbox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                        textbox.TextFrame.TextRange.Font.Size = codeTextBeforeEdit.Font.Size;
                        textbox.TextFrame.TextRange.Font.Name = codeTextBeforeEdit.Font.Name;
                        textbox.TextEffect.Alignment = Office.MsoTextEffectAlignment.msoTextEffectAlignmentLeft;

                        // Add the created textbox reference to a list for further processing
                        transitionText.Add(Tuple.Create(diffs[j].Type, textbox));

                        // Syntax Highlighting for textbox
                        for (int charIndex = 1; charIndex <= lines[i].Length; charIndex++)
                        {
                            if (diffs[j].Type == ChangeType.Deleted || diffs[j].Type == ChangeType.Unchanged)
                            {
                                textbox.TextFrame.TextRange.Characters(charIndex, 1).Font.Color.RGB = codeTextBeforeEdit.Characters(charCountBefore, 1).Font.Color.RGB;
                                charCountBefore++;
                            }
                        }

                        // Increment the syntax highlighter pointer to accommodate new line
                        if (i < lines.Length - 1 && (diffs[j].Type == ChangeType.Deleted || diffs[j].Type == ChangeType.Unchanged))
                        {
                            charCountBefore++;
                        }
                    }

                    // Set left pointer to end of newly created textbox
                    leftPointer += textbox.Width - (2 * emptyTextboxOffset);

                    // Increment the pointer for code lines containing only Equal lines if line is equal
                    if (diffs[j].Type != ChangeType.Deleted)
                    {
                        leftEqualPointer += textbox.Width - (2 * emptyTextboxOffset);
                    }

                    // Set pointers to next line if the code line ends with a newline
                    if (lines[lines.Length - 1].Contains("\n"))
                    {
                        leftPointer = originalLeftPointer;
                        leftEqualPointer = originalLeftPointer;
                        topPointer += topPointerLineOffset;
                    }

                    // Increment the syntax highlighter pointer if there is a new line
                    if (diffs[j].Type != ChangeType.Inserted && charCountBefore <= codeTextBeforeEdit.Characters().Length && Char.IsControl(codeTextBeforeEdit.Characters(charCountBefore, 1).Text, 0))
                    {
                        charCountBefore++;
                    }

                }

                // Update all left and top pointers according to their Diff types
                if (diffs[j].Type == ChangeType.Deleted)
                {
                    leftBeforePointer = leftPointer;
                    topBeforePointer = topPointer;
                }
                else if (diffs[j].Type == ChangeType.Inserted)
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

            // Remove all empty textboxes created
            for (int i = transitionText.Count - 1; i >= 0; i--)
            {
                if (transitionText[i].Item2.TextFrame.TextRange.Length == 0)
                {
                    transitionText.RemoveAt(i);
                }
            }

            // Creates tuples that stores the diff types of successive pairs of text boxes for animation purposes
            for (int i = 1; i < transitionText.Count; i++)
            {
                WordDiffType transitionTextType;
                if (transitionText[i - 1].Item1 == ChangeType.Deleted && transitionText[i].Item1 == ChangeType.Inserted)
                {
                    transitionTextType = WordDiffType.DeleteAdd;
                }
                else if (transitionText[i - 1].Item1 == ChangeType.Deleted)
                {
                    transitionTextType = WordDiffType.DeleteEqual;
                }
                else if (transitionText[i - 1].Item1 == ChangeType.Inserted)
                {
                    transitionTextType = WordDiffType.AddEqual;
                }
                else
                {
                    transitionTextType = WordDiffType.Equal;
                }
                if (transitionTextType != WordDiffType.Equal)
                {
                    transitionTextToAnimate.Add(Tuple.Create(transitionTextType, transitionText[i - 1].Item2, transitionText[i].Item2));
                }
                if (transitionTextType == WordDiffType.DeleteAdd)
                {
                    i++;
                }
            }

            // Create a tuple that stores the diff type of the last pair of text boxes for animation purposes
            if (transitionText[transitionText.Count - 1].Item1 == ChangeType.Deleted)
            {
                transitionTextToAnimate.Add(Tuple.Create(WordDiffType.DeleteEqual, transitionText[transitionText.Count - 1].Item2, transitionText[transitionText.Count - 1].Item2));
            }
            else if (transitionText[transitionText.Count - 1].Item1 == ChangeType.Inserted)
            {
                transitionTextToAnimate.Add(Tuple.Create(WordDiffType.AddEqual, transitionText[transitionText.Count - 1].Item2, transitionText[transitionText.Count - 1].Item2));
            }
            return transitionTextToAnimate;
        }
        private void BuildDiffPieces(DiffPlex.Model.DiffResult diffResult, List<DiffPiece> pieces)
        {
            int bPos = 0;

            foreach (var diffBlock in diffResult.DiffBlocks)
            {
                for (; bPos < diffBlock.InsertStartB; bPos++)
                {
                    pieces.Add(new DiffPiece(diffResult.PiecesNew[bPos], ChangeType.Unchanged, bPos + 1));
                }

                int i = 0;
                for (; i < Math.Min(diffBlock.DeleteCountA, diffBlock.InsertCountB); i++)
                {
                    pieces.Add(new DiffPiece(diffResult.PiecesOld[i + diffBlock.DeleteStartA], ChangeType.Deleted));
                }

                i = 0;
                for (; i < Math.Min(diffBlock.DeleteCountA, diffBlock.InsertCountB); i++)
                {
                    pieces.Add(new DiffPiece(diffResult.PiecesNew[i + diffBlock.InsertStartB], ChangeType.Inserted, bPos + 1));
                    bPos++;
                }

                if (diffBlock.DeleteCountA > diffBlock.InsertCountB)
                {
                    for (; i < diffBlock.DeleteCountA; i++)
                    {
                        pieces.Add(new DiffPiece(diffResult.PiecesOld[i + diffBlock.DeleteStartA], ChangeType.Deleted));
                    }
                }
                else
                {
                    for (; i < diffBlock.InsertCountB; i++)
                    {
                        pieces.Add(new DiffPiece(diffResult.PiecesNew[i + diffBlock.InsertStartB], ChangeType.Inserted, bPos + 1));
                        bPos++;
                    }
                }
            }

            for (; bPos < diffResult.PiecesNew.Length; bPos++)
            {
                pieces.Add(new DiffPiece(diffResult.PiecesNew[bPos], ChangeType.Unchanged, bPos + 1));
            }
        }

        /// <summary>
        /// Apply formatting and timings to delete effects in word diff animation to simulate code deletion
        /// </summary>
        /// <param name="effectList">list of effects to format</param>
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

        /// <summary>
        /// Apply formatting and timings to add effects in word diff animation to simulate code addition
        /// </summary>
        /// <param name="effectList">list of effects to format</param>
        private void FormatWordDiffAddEffects(List<PowerPoint.Effect> effectList)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                effect.Timing.Duration = 0.7f;
                effect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }
        }

        /// <summary>
        /// Apply formatting and timings to colour change effects in word diff animation to simulate highlighting of code to be modified.
        /// </summary>
        /// <param name="effectList">list of effects to format</param>
        private void FormatWordDiffColourChangeEffects(List<PowerPoint.Effect> effectList)
        {
            foreach (PowerPoint.Effect effect in effectList)
            {
                effect.Timing.Duration = 0.1f;
                effect.EffectParameters.Color2.RGB = Utils.GraphicsUtil.ConvertColorToRgb(LiveCodingLabSettings.bulletsTextHighlightColor);
            }
        }

        /// <summary>
        /// Apply formatting and timings to move left effects in word diff animation to simulate code moving left to close up gaps from deletion.
        /// </summary>
        /// <param name="effectList">list of effects to format</param>
        /// <param name="offset">distance for code to shift left by</param>
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

        /// <summary>
        /// Apply formatting and timings to move right effects in word diff animation to simulate code moving right to create space for addition.
        /// </summary>
        /// <param name="effectList">list of effects to format</param>
        /// <param name="offset">distance for code to shift right by</param>
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
    }
}
