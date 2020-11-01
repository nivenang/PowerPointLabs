using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.LiveCodingLab.Views
{
#pragma warning disable 0618
    /// <summary>
    /// Interaction logic for LiveCodingPaneWPF.xaml
    /// </summary>
    public partial class LiveCodingPaneWPF : UserControl, ILiveCodingLabPane
    {
        private LiveCodingLabMain _liveCodingLab;
        private readonly LiveCodingLabErrorHandler _errorHandler;
        private ObservableCollection<CodeBoxPaneItem> codeBoxList;
        private PowerPointPresentation currentPresentation;
        private CollectionView view;
        private PropertyGroupDescription groupDescription;

        #region Interface Implementation
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                ErrorDialogBox.ShowDialog(TextCollection.CommonText.ErrorTitle, content, exception);
            }
            else
            {
                MessageBox.Show(content, TextCollection.CommonText.ErrorTitle);
            }
        }

        public void Reset()
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes(false);

            if (selectedShapes != null)
            {
                this.ExecuteOfficeCommand("Undo");
                GC.Collect();
            }
        }

        public void ExecuteLiveCodingAction(List<CodeBoxPaneItem> listCodeBox, Action<List<CodeBoxPaneItem>> liveCodingAction)
        {
            if (listCodeBox == null || listCodeBox.Count != 2)
            {
                return;
            }

            this.StartNewUndoEntry();
            liveCodingAction.Invoke(listCodeBox);
        }
        public void ExecuteLiveCodingAction(string diffPath, Action<string, LiveCodingPaneWPF, string> liveCodingAction, string diffGroupName)
        {
            if (diffPath == null || diffPath.Trim() == "")
            {
                return;
            }

            this.StartNewUndoEntry();
            liveCodingAction.Invoke(diffPath, this, diffGroupName);
        }

        #endregion
        public LiveCodingPaneWPF()
        {
            InitializeComponent();
            InitialiseLogic();
            currentPresentation = PowerPointPresentation.Current;
            _errorHandler = LiveCodingLabErrorHandler.InitializeErrorHandler(this);
            codeBoxList = LoadCodeBoxes(currentPresentation.FirstSlide);
            Focusable = true;
            codeListBox.ItemsSource = codeBoxList;
            view = (CollectionView)CollectionViewSource.GetDefaultView(codeListBox.ItemsSource);
            groupDescription = new PropertyGroupDescription("Group");
            view.GroupDescriptions.Add(groupDescription);
            RefreshCode();
        }
        public void RemoveCodeBox(Object codeBox)
        {
            int index = 0;
            while (index < codeBoxList.Count)
            {
                if (codeBoxList[index] == codeBox)
                {
                    codeBoxList.RemoveAt(index);
                }
                else
                {
                    index++;
                }
            }
        }

        public void AddCodeBox(CodeBoxPaneItem item)
        {
            codeBoxList.Insert(0, item);
            codeListBox.SelectedIndex = 0;
            SaveCodeBox();
        }

        public void SaveCodeBox()
        {
            LiveCodingLabTextStorageService.StoreCodeBoxToSlide(codeBoxList, currentPresentation.FirstSlide);
        }

        public void MoveUpCodeBox(CodeBoxPaneItem item)
        {
            for (int index = 0; index < codeBoxList.Count; index++)
            {
                if (codeBoxList[index] == item && index == 0)
                {
                    return;
                }

                if (codeBoxList[index] == item && index > 0)
                {
                    for (int i = index-1; i >= 0; i--)
                    {
                        if (codeBoxList[index].Group == codeBoxList[i].Group)
                        {
                            CodeBoxPaneItem temp = codeBoxList[index];
                            codeBoxList[index] = codeBoxList[i];
                            codeBoxList[i] = temp;
                            break;
                        }
                    }

                    break;
                }
            }
        }

        public void MoveDownCodeBox(CodeBoxPaneItem item)
        {
            for (int index = 0; index < codeBoxList.Count; index++)
            {
                if (codeBoxList[index] == item && index == codeBoxList.Count - 1)
                {
                    return;
                }

                if (codeBoxList[index] == item && index < codeBoxList.Count - 1)
                {
                    for (int i = index + 1; i < codeBoxList.Count; i++)
                    {
                        if (codeBoxList[index].Group == codeBoxList[i].Group)
                        {
                            CodeBoxPaneItem temp = codeBoxList[index];
                            codeBoxList[index] = codeBoxList[i];
                            codeBoxList[i] = temp;
                            break;
                        }
                    }

                    break;
                }
            }
        }

        public void RefreshCode()
        {
            foreach (PowerPointSlide slide in currentPresentation.Slides)
            {
                if (slide.Name.Contains(LiveCodingLabText.TransitionSlideIdentifier))
                {
                    continue;
                }
                List<Shape> shapes = slide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);
                foreach (Shape shape in shapes)
                {
                    CodeBoxPaneItem codeBox = GetCodeBoxPaneItemFromShape(shape);
                    if (codeBox != null)
                    {
                        CodeBox code = codeBox.CodeBox;
                        code.Slide = slide;
                        code.Shape = shape;
                        codeBox.CodeBox = code;
                    }
                }
            }
            foreach (CodeBoxPaneItem item in codeBoxList)
            {
                PowerPointSlide slide = item.CodeBox.Slide;
                if (slide == null)
                {
                    continue;
                }
                try
                {
                    int slideID = slide.ID;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    item.CodeBox.Slide = null;
                    item.CodeBox.Shape = null;
                }
            }

            SaveCodeBox();
        }
        internal void InitialiseLogic()
        {
            if (_liveCodingLab == null)
            {
                _liveCodingLab = new LiveCodingLabMain();
            }
        }

        #region API
        private CodeBoxPaneItem AddCodeBoxToList()
        {
            CodeBoxPaneItem item = new CodeBoxPaneItem(this);
            codeBoxList.Insert(0, item);
            codeListBox.SelectedIndex = 0;
            return item;
        }

        #endregion

        #region XAML-Binded Event Handler

        private void InsertCodeBoxButton_Click(object sender, RoutedEventArgs e)
        {
            AddCodeBoxToList();
            SaveCodeBox();
        }

        private void RefreshCodeButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            foreach (CodeBoxPaneItem item in codeListBox.SelectedItems)
            {
                if (item != null)
                {
                    item.CodeBox.Text = item.codeTextBox.Text;
                    if (item.CodeBox.Shape != null)
                    {
                        item.CodeBox = ShapeUtility.ReplaceTextForShape(item.CodeBox);
                    }
                }
            }
            SaveCodeBox();
        }

        private void RefreshAllCodeButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            foreach (CodeBoxPaneItem item in codeListBox.Items)
            {
                if (item != null)
                {
                    item.CodeBox.Text = item.codeTextBox.Text;
                    if (item.CodeBox.Shape != null)
                    {
                        item.CodeBox = ShapeUtility.ReplaceTextForShape(item.CodeBox);
                    }
                }
            }
            SaveCodeBox();
        }

        private void GroupCodeButton_Click(object sender, RoutedEventArgs e)
        {
            GroupCodeBoxDialog dialog = new GroupCodeBoxDialog();
            string defaultGroupName = "";
            if (dialog.ShowThematicDialog() == true)
            {
                defaultGroupName = dialog.ResponseText;
            }
            foreach (CodeBoxPaneItem item in codeListBox.SelectedItems)
            {
                if (item != null && defaultGroupName != "")
                {
                    item.Group = defaultGroupName;
                }
            }
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(groupDescription);
            SaveCodeBox();
        }

        private void InsertDiffButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            InsertDiffDialog diffDialog = new InsertDiffDialog();
            string diffPath = "";
            string diffGroup = "Ungrouped";
            if (diffDialog.ShowThematicDialog() == true)
            {
                diffPath = diffDialog.DiffFile;
                if (diffDialog.DiffGroup != "")
                {
                    diffGroup = diffDialog.DiffGroup;
                }
            }
            if (diffPath == "")
            {
                return;
            }

            Action<string, LiveCodingPaneWPF, string> insertDiffAction = (diffFilePath, liveCodingPane, diffGroupName) => _liveCodingLab.InsertDiff(diffFilePath, liveCodingPane, diffGroupName);
            ClickHandler(insertDiffAction, diffPath, diffGroup);
        }

        private void HighlightDifferenceButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            Action<List<CodeBoxPaneItem>> highlightDifferenceAction = codeBoxes => _liveCodingLab.HighlightDifferences(codeBoxes);
            ClickHandler(highlightDifferenceAction, 1, LiveCodingLabMain.HighlightDifference_ErrorParameters);
        }

        private void AnimateNewLinesButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            Action<List<CodeBoxPaneItem>> animateNewLinesAction = codeBoxes => _liveCodingLab.AnimateNewLines(codeBoxes);
            ClickHandler(animateNewLinesAction, 1, LiveCodingLabMain.AnimateNewLines_ErrorParameters);
        }
        private void AnimateLineDiffButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            Action<List<CodeBoxPaneItem>> animateLineDiffAction = codeBoxes => _liveCodingLab.AnimateLineDiff(codeBoxes);
            ClickHandler(animateLineDiffAction, LiveCodingLabMain.AnimateLineDiff_ErrorParameters);
        }
        private void AnimateBlockDiffButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshCode();
            Action<List<CodeBoxPaneItem>> animateBlockDiffAction = codeBoxes => _liveCodingLab.AnimateBlockDiff(codeBoxes);
            ClickHandler(animateBlockDiffAction, LiveCodingLabMain.AnimateBlockDiff_ErrorParameters);
        }
        private void AnimationSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            LiveCodingLabSettings.ShowAnimationSettingsDialog();
        }

        private void CodeBoxSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            LiveCodingLabSettings.ShowCodeBoxSettingsDialog();
        }
        #endregion

        #region Helper Methods

        private ObservableCollection<CodeBoxPaneItem> LoadCodeBoxes(PowerPointSlide slide)
        {
            ObservableCollection<CodeBoxPaneItem> codeBoxesList = new ObservableCollection<CodeBoxPaneItem>();
            List<Dictionary<string, string>> codeBoxes =
                LiveCodingLabTextStorageService.LoadCodeBoxesFromSlide(slide);

            while (codeBoxes != null && codeBoxes.Count > 0)
            {
                CodeBoxPaneItem codeBox = CreateCodeBoxFromDictionary(codeBoxes.First());
                codeBoxesList.Add(codeBox);
                codeBoxes.RemoveAt(0);
            }

            return codeBoxesList;
        }

        private CodeBoxPaneItem CreateCodeBoxFromDictionary(Dictionary<string, string> codeBoxItemDic)
        {
            int id = int.Parse(codeBoxItemDic[LiveCodingLabText.CodeBox_Id]);
            CodeBoxIdService.PopulateCodeBoxIds(id);
            bool isFile = codeBoxItemDic[LiveCodingLabText.CodeBox_IsFile] == "Y";
            bool isText = codeBoxItemDic[LiveCodingLabText.CodeBox_IsText] == "Y";
            bool isDiff = codeBoxItemDic[LiveCodingLabText.CodeBox_IsDiff] == "Y";
            int diffIndex = int.Parse(codeBoxItemDic[LiveCodingLabText.CodeBox_DiffIndex]);
            string group = codeBoxItemDic[LiveCodingLabText.CodeBox_Group];
            string shapeName = codeBoxItemDic[LiveCodingLabText.CodeBox_ShapeName];
            PowerPointSlide slide = null;
            CodeBox codeBoxItem;

            if (isFile)
            {
                codeBoxItem = new CodeBox(id,
                    codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], "", "", isFile, false, false, slide, shapeName);
            }
            else if (isText)
            {
                codeBoxItem = new CodeBox(id,
                    "", codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], "", false, isText, false, slide, shapeName);
            }
            else
            {
                codeBoxItem = new CodeBox(id,
                    "", "", codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], false, false, isDiff, slide, shapeName, diffIndex);
            }

            CodeBoxPaneItem codeBoxPaneItem = new CodeBoxPaneItem(this, codeBoxItem, group);
            return codeBoxPaneItem;
        }

        private PowerPoint.ShapeRange GetSelectedShapes(bool handleError = false)
        {
            PowerPoint.Selection selection = this.GetCurrentSelection();
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                return null;
            }
            else if (selection.ShapeRange.Count > 1)
            {
                return null;
            }
            else
            {
                return selection.ShapeRange;
            }
        }
        private void ClickHandler(Action<List<CodeBoxPaneItem>> liveCodingAction, int minNoOfSelectedShapes, string[] errorParameters)
        {
            List<CodeBoxPaneItem> listCodeBox = new List<CodeBoxPaneItem>();
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }
            PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];
            List<PowerPoint.Shape> shapesToUseCurrentSlide = currentSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);
            List<PowerPoint.Shape> shapesToUseNextSlide = nextSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);

            if (shapesToUseCurrentSlide == null || shapesToUseNextSlide == null)
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }

            if (shapesToUseCurrentSlide.Count != 1 || !HasText(shapesToUseCurrentSlide[0]))
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }

            if (shapesToUseNextSlide.Count != 1 || !HasText(shapesToUseNextSlide[0]))
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }
            foreach (CodeBoxPaneItem item in codeListBox.Items)
            {
                if (item != null && item.CodeBox.Shape != null && (item.CodeBox.Shape.Name == shapesToUseCurrentSlide[0].Name))
                {
                    listCodeBox.Add(item);
                    break;
                }
            }

            foreach (CodeBoxPaneItem item in codeListBox.Items)
            {
                if (item != null && item.CodeBox.Shape != null && (item.CodeBox.Shape.Name == shapesToUseNextSlide[0].Name))
                {
                    listCodeBox.Add(item);
                    break;
                }
            }
            ExecuteLiveCodingAction(listCodeBox, liveCodingAction);
        }
        private void ClickHandler(Action<List<CodeBoxPaneItem>> liveCodingAction, string[] errorParameters)
        {
            List<CodeBoxPaneItem> listCodeBox = new List<CodeBoxPaneItem>();
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }
            PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];

            List<PowerPoint.Shape> shapesToUseCurrentSlide = currentSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);
            List<PowerPoint.Shape> shapesToUseNextSlide = nextSlide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);

            if (shapesToUseCurrentSlide == null || shapesToUseNextSlide == null)
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }

            if (shapesToUseCurrentSlide.Count != 1 || !HasText(shapesToUseCurrentSlide[0]))
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }

            if (shapesToUseNextSlide.Count != 1 || !HasText(shapesToUseNextSlide[0]))
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                return;
            }
            foreach (CodeBoxPaneItem item in codeListBox.Items)
            {
                if (item != null && item.CodeBox.Shape != null && (item.CodeBox.Shape.Name == shapesToUseCurrentSlide[0].Name))
                {
                    listCodeBox.Add(item);
                    break;
                }
            }

            foreach (CodeBoxPaneItem item in codeListBox.Items)
            {
                if (item != null && item.CodeBox.Shape != null && (item.CodeBox.Shape.Name == shapesToUseNextSlide[0].Name))
                {
                    listCodeBox.Add(item);
                    break;
                }
            }

            foreach (CodeBoxPaneItem item in listCodeBox)
            {
                if (!item.CodeBox.IsDiff)
                {
                    _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidCodeBox, errorParameters);
                    return;
                }
            }
            ExecuteLiveCodingAction(listCodeBox, liveCodingAction);
        }

        private void ClickHandler(Action<string, LiveCodingPaneWPF, string> liveCodingAction, string diffPath, string diffGroupName)
        {
            ExecuteLiveCodingAction(diffPath, liveCodingAction, diffGroupName);
        }
        private CodeBoxPaneItem GetCodeBoxPaneItemFromShape(Shape shape)
        {
            foreach (CodeBoxPaneItem codeBox in codeBoxList)
            {
                if (codeBox.CodeBox.ShapeName == shape.Name)
                {
                    return codeBox;
                }
            }
            return null;
        }

        /// <summary>
        /// Returns true iff shape has a text frame.
        /// </summary>
        private static bool HasText(Shape shape)
        {
            return shape.HasTextFrame == MsoTriState.msoTrue &&
                   shape.TextFrame2.HasText == MsoTriState.msoTrue;
        }
        #endregion
    }
}
