using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DiffPlex.DiffBuilder;
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

                // Check that there is a slide selected by the user
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }

                // Check that there exists a "before" and "after" code
                if (codeListBox.Count != 2)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                CodeBoxPaneItem diffCodeBoxBefore = codeListBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = codeListBox[1];

                FileDiff diffFile;

                // Case 1: Animating differences across a Diff File
                if (diffCodeBoxBefore.CodeBox.IsDiff && diffCodeBoxAfter.CodeBox.IsDiff)
                {
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

                    diffFile = diffList[0];
                }
                // Case 2: Animating differences across two user-input code snippets by building a diff file
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

                    var diff = InlineDiffBuilder.Diff(diffCodeBoxBefore.CodeBox.Text, diffCodeBoxAfter.CodeBox.Text);
                    diffFile = BuildDiffFromText(diffCodeBoxBefore.CodeBox.Text, diffCodeBoxAfter.CodeBox.Text);
                }
                // Default: Inform user that code snippets to be animated do not match up
                else
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                AnimateDiff(codeListBox, diffFile, true);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateBlockDiff");
                throw;
            }
        }
    }
}
