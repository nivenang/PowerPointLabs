using System.Collections.Generic;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using TestInterface;

using Point = System.Drawing.Point;

namespace Test.FunctionalTest
{
    [TestClass]
    public class LiveCodingLabTest : BaseFunctionalTest
    {
        private const string InsertTextCodeBoxShape = "InsertTextCodeBoxTestShape";
        private const string InsertFileCodeBoxShape = "InsertFileCodeBoxTestShape";
        private const string RefreshCodeBoxShape = "RefreshCodeBoxTestShape";
        private const string InsertDiffCodeBoxShape = "PPTL";
        private const string InsertDiffCodeBoxBeforeOriginalShape = "InsertDiffCodeBeforeOriginalShape";
        private const string InsertDiffCodeBoxAfterOriginalShape = "InsertDiffCodeAfterOriginalShape";

        private const string InsertCodeBoxText = "Insert Text Code Box Test\n";
        private const string InsertCodeBoxFile = "LiveCodingLab\\sample.txt";
        private const string InsertCodeBoxFileText = "Insert File Code Box Test\n";
        private const string InsertCodeBoxDiff = "LiveCodingLab\\sample.diff";
        private const string RefreshCodeBoxTextBefore = "Before Refresh Code Text\n";
        private const string RefreshCodeBoxTextAfter = "After Refresh Code Text\n";

        private const int TestInsertTextCodeBoxSlideNo = 3;
        private const int TestInsertFileCodeBoxSlideNo = 4;
        private const int TestRefreshCodeSlideNo = 5;
        private const int TestInsertDiffOriginalShapeSlideNo = 6;
        private const int TestInsertDiffBeforeCodeBoxSlideNo = 7;
        private const int TestInsertDiffAfterCodeBoxSlideNo = 8;
        private const int TestAnimateLineDiffExpectedSlideNo = 10;
        private const int TestAnimateLineDiffSlideNo = 11;
        private const int TestAnimateLineDiffActualSlideNo = 12;
        private const int TestAnimateBlockDiffExpectedSlideNo = 14;
        private const int TestAnimateBlockDiffSlideNo = 15;
        private const int TestAnimateBlockDiffActualSlideNo = 16;
        private const int TestAnimateWordDiffExpectedSlideNo = 18;
        private const int TestAnimateWordDiffSlideNo = 19;
        private const int TestAnimateWordDiffActualSlideNo = 20;

        protected override string GetTestingSlideName()
        {
            return "LiveCodingLab\\LiveCodingLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_LiveCodingLabTest()
        {
            PpOperations.MaximizeWindow();
            ILiveCodingLabController liveCodingLab = PplFeatures.LiveCodingLab;
            liveCodingLab.OpenPane();
            ThreadUtil.WaitFor(1000);
            
            // Code Box Tests
            TestInsertTextCodeBox(liveCodingLab);
            TestInsertFileCodeBox(liveCodingLab);
            TestRefreshCode(liveCodingLab);
            TestInsertDiffCodeBox(liveCodingLab);

            // Animation Tests
            TestAnimateLineDiff(liveCodingLab);
            TestAnimateBlockDiff(liveCodingLab);
        }

        #region Code Box Tests
        private void TestInsertTextCodeBox(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestInsertTextCodeBoxSlideNo);

            liveCodingLab.InsertTextCodeBox(InsertCodeBoxText, InsertTextCodeBoxShape);
            
            ThreadUtil.WaitFor(1000);
            Assert.AreEqual(PpOperations.SelectAllTextInShape(InsertTextCodeBoxShape).Trim(), InsertCodeBoxText.Trim());
        }

        private void TestInsertFileCodeBox(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestInsertFileCodeBoxSlideNo);

            liveCodingLab.InsertFileCodeBox(PathUtil.GetDocTestPath() + InsertCodeBoxFile);

            ThreadUtil.WaitFor(1000);
            Assert.AreEqual(PpOperations.SelectAllTextInShape(InsertFileCodeBoxShape).Trim(), InsertCodeBoxFileText.Trim());
        }

        private void TestRefreshCode(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestRefreshCodeSlideNo);

            liveCodingLab.InsertTextCodeBox(RefreshCodeBoxTextBefore, RefreshCodeBoxShape);

            ThreadUtil.WaitFor(1000);

            Assert.AreEqual(PpOperations.SelectAllTextInShape(RefreshCodeBoxShape).Trim(), RefreshCodeBoxTextBefore.Trim());

            liveCodingLab.RefreshTextCodeBox(RefreshCodeBoxTextBefore, RefreshCodeBoxTextAfter);

            ThreadUtil.WaitFor(1000);

            Assert.AreEqual(PpOperations.SelectAllTextInShape(RefreshCodeBoxShape).Trim(), RefreshCodeBoxTextAfter.Trim());
        }

        private void TestInsertDiffCodeBox(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestInsertDiffOriginalShapeSlideNo);

            string InsertDiffBeforeText = PpOperations.SelectAllTextInShape(InsertDiffCodeBoxBeforeOriginalShape).Trim();
            string InsertDiffAfterText = PpOperations.SelectAllTextInShape(InsertDiffCodeBoxAfterOriginalShape).Trim();

            liveCodingLab.InsertDiffCodeBox(PathUtil.GetDocTestPath() + InsertCodeBoxDiff);

            ThreadUtil.WaitFor(1000);

            PpOperations.SelectSlide(TestInsertDiffBeforeCodeBoxSlideNo);
            Assert.AreEqual(PpOperations.SelectShapesByPrefix(InsertDiffCodeBoxShape).TextFrame.TextRange.Text.Trim(),
                InsertDiffBeforeText);
            
            PpOperations.SelectSlide(TestInsertDiffAfterCodeBoxSlideNo);
            Assert.AreEqual(PpOperations.SelectShapesByPrefix(InsertDiffCodeBoxShape).TextFrame.TextRange.Text.Trim(),
                InsertDiffAfterText);
        }
        #endregion

        #region Animation Tests
        private void TestAnimateLineDiff(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestAnimateLineDiffSlideNo);

            liveCodingLab.AnimateLineDiff();

            Slide expectedSlide = PpOperations.SelectSlide(TestAnimateLineDiffExpectedSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(TestAnimateLineDiffActualSlideNo);

            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
        }

        private void TestAnimateBlockDiff(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestAnimateBlockDiffSlideNo);

            liveCodingLab.AnimateBlockDiff();

            Slide expectedSlide = PpOperations.SelectSlide(TestAnimateBlockDiffExpectedSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(TestAnimateBlockDiffActualSlideNo);

            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
        }

        /*
        private void TestAnimateWordDiff(ILiveCodingLabController liveCodingLab)
        {
            PpOperations.SelectSlide(TestAnimateWordDiffSlideNo);

            liveCodingLab.AnimateWordDiff();

            Slide expectedSlide = PpOperations.SelectSlide(TestAnimateWordDiffExpectedSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(TestAnimateWordDiffActualSlideNo);

            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
        }
        */
        #endregion

        #region Helper methods
        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown(startPt.X, startPt.Y);
            MouseUtil.SendMouseUp(endPt.X, endPt.Y);
        }

        private void Click(Control target)
        {
            Point pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
        }
        # endregion
    }
}
