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
        public void InsertDiff(string diffPath, LiveCodingPaneWPF parent, string diffGroup)
        {
            try
            {
                // Parses the user-input diff file 
                List<FileDiff> diffList = CodeBoxFileService.ParseDiff(diffPath);

                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                // Check that the user has selected a slide
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }
                
                // Check that there is at least one parsed diff block in the file
                if (diffList.Count < 1) 
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                int currentSlideIndex = currentSlide.Index;
                foreach (FileDiff diff in diffList) 
                {
                    // Create new slides to insert the "before" and "after" code
                    PowerPointSlide diffSlideBefore = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlideIndex + 1);
                    PowerPointSlide diffSlideAfter = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlideIndex + 2);

                    // Creates a Code Box object for both the "before" and "after" code snippets
                    CodeBoxPaneItem codeBoxPaneItemBefore = new CodeBoxPaneItem(parent);
                    CodeBoxPaneItem codeBoxPaneItemAfter = new CodeBoxPaneItem(parent);

                    codeBoxPaneItemBefore.SetDiff();
                    codeBoxPaneItemAfter.SetDiff();
                    codeBoxPaneItemBefore.CodeBox.Text = diffPath;
                    codeBoxPaneItemBefore.Group = diffGroup;
                    codeBoxPaneItemAfter.CodeBox.Text = diffPath;
                    codeBoxPaneItemAfter.Group = diffGroup;
                    codeBoxPaneItemBefore.CodeBox.DiffIndex = 0;
                    codeBoxPaneItemAfter.CodeBox.DiffIndex = 1;

                    // Add new Code Box objects to the ObservableCollection so users can see on the Live Coding Pane
                    parent.AddCodeBox(codeBoxPaneItemBefore);
                    parent.AddCodeBox(codeBoxPaneItemAfter);

                    // Insert Code Box objects as text box into their respective slides
                    CodeBox diffCodeBoxBefore = ShapeUtility.InsertDiffCodeBoxToSlide(diffSlideBefore, codeBoxPaneItemBefore.CodeBox, diff);
                    CodeBox diffCodeBoxAfter = ShapeUtility.InsertDiffCodeBoxToSlide(diffSlideAfter, codeBoxPaneItemAfter.CodeBox, diff);

                    codeBoxPaneItemBefore.CodeBox = diffCodeBoxBefore;
                    codeBoxPaneItemAfter.CodeBox = diffCodeBoxAfter;
                    currentSlideIndex += 2;
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
    }
}
