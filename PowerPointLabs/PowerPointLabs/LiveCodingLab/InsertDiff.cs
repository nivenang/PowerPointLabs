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
                List<FileDiff> diffList = CodeBoxFileService.ParseDiff(diffPath);

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
                int currentSlideIndex = currentSlide.Index;
                foreach (FileDiff diff in diffList) 
                {
                    PowerPointSlide diffSlideBefore = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlideIndex + 1);
                    PowerPointSlide diffSlideAfter = currentPresentation.AddSlide(PowerPoint.PpSlideLayout.ppLayoutOrgchart, index: currentSlideIndex + 2);
                    // INSERT LOGIC FOR INSERTING SHAPE INTO SLIDE
                    // CREATE CODEBOXPANEITEM FROM DIFF

                    CodeBoxPaneItem codeBoxPaneItemBefore = new CodeBoxPaneItem(parent);
                    CodeBoxPaneItem codeBoxPaneItemAfter = new CodeBoxPaneItem(parent);
                    // UPDATE CODEBOX TO HAVE NEW TEXT
                    codeBoxPaneItemBefore.SetDiff();
                    codeBoxPaneItemAfter.SetDiff();
                    codeBoxPaneItemBefore.CodeBox.Text = diffPath;
                    codeBoxPaneItemBefore.Group = diffGroup;
                    codeBoxPaneItemAfter.CodeBox.Text = diffPath;
                    codeBoxPaneItemAfter.Group = diffGroup;
                    codeBoxPaneItemBefore.CodeBox.DiffIndex = 0;
                    codeBoxPaneItemAfter.CodeBox.DiffIndex = 1;

                    parent.AddCodeBox(codeBoxPaneItemBefore);
                    parent.AddCodeBox(codeBoxPaneItemAfter);

                    CodeBox diffCodeBoxBefore = ShapeUtility.InsertDiffCodeBoxToSlide(diffSlideBefore, codeBoxPaneItemBefore.CodeBox, diff);
                    // INSERT LOGIC FOR INSERTING SHAPE INTO SLIDE
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
