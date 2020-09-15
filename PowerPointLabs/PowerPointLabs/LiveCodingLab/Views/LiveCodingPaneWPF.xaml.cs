using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
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
        private List<CodeBoxPaneItem> codeBoxList;
        private PowerPointPresentation currentPresentation;

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

        public void ExecuteLiveCodingAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> liveCodingAction)
        {
            if (selectedShapes == null)
            {
                return;
            }

            this.StartNewUndoEntry();
            liveCodingAction.Invoke(selectedShapes);
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
        }
        public void RemoveCodeBox(Object codeBox)
        {
            int index = 0;
            while (index < codeListBox.Items.Count)
            {
                if (codeListBox.Items[index] == codeBox)
                {
                    codeListBox.Items.RemoveAt(index);
                    codeBoxList.RemoveAt(index);
                }
                else
                {
                    index++;
                }
            }
        }

        public void SaveCodeBox()
        {
            LiveCodingLabTextStorageService.StoreCodeBoxToSlide(codeBoxList, currentPresentation.FirstSlide);
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
            codeListBox.Items.Insert(0, item);
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
            SaveCodeBox();

        }

        private void RefreshAllCodeButton_Click(object sender, RoutedEventArgs e)
        {
            PopulateCodeBoxPaneItemTextBoxes();
            SaveCodeBox();
        }

        private void HighlightDifferenceButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> highlightDifferenceAction = shapes => _liveCodingLab.HighlightDifferences(shapes);
            ClickHandler(highlightDifferenceAction, 1, LiveCodingLabMain.HighlightDifference_ErrorParameters);
        }

        private void AnimateNewLinesButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> animateNewLinesAction = shapes => _liveCodingLab.AnimateNewLines(shapes);
            ClickHandler(animateNewLinesAction, 1, LiveCodingLabMain.AnimateNewLines_ErrorParameters);
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

        private void PopulateCodeBoxPaneItemTextBoxes()
        {
            foreach (CodeBoxPaneItem item in codeBoxList)
            {
                item.PopulateTextBox();
            }
        }

        private List<CodeBoxPaneItem> LoadCodeBoxes(PowerPointSlide slide)
        {
            List<CodeBoxPaneItem> codeBoxesList = new List<CodeBoxPaneItem>();
            List<Dictionary<string, string>> codeBoxes =
                LiveCodingLabTextStorageService.LoadCodeBoxesFromSlide(slide);

            while (codeBoxes != null && codeBoxes.Count > 0)
            {
                CodeBoxPaneItem codeBox = CreateCodeBoxFromDictionary(codeBoxes.First());
                codeBoxesList.Add(codeBox);
                codeBoxes.RemoveAt(0);
                codeListBox.Items.Insert(0, codeBox);
                codeListBox.SelectedIndex = 0;
            }
            return codeBoxesList;
        }

        private CodeBoxPaneItem CreateCodeBoxFromDictionary(Dictionary<string, string> codeBoxItemDic)
        {
            int id = int.Parse(codeBoxItemDic[LiveCodingLabText.CodeBox_Id]);
            CodeBoxIdService.PopulateCodeBoxIds(id);
            bool isURL = codeBoxItemDic[LiveCodingLabText.CodeBox_IsURL] == "Y";
            bool isFile = codeBoxItemDic[LiveCodingLabText.CodeBox_IsFile] == "Y";
            bool isText = codeBoxItemDic[LiveCodingLabText.CodeBox_IsText] == "Y";
            int slideNum = int.Parse(codeBoxItemDic[LiveCodingLabText.CodeBox_Slide]);
            PowerPointSlide slide = null;
            CodeBox codeBoxItem;
            Shape codeShape = null;

            if (slideNum > 0)
            {
                slide = PowerPointPresentation.Current.GetSlide(slideNum);
                List<Shape> shapes = slide.GetShapesWithNameRegex(LiveCodingLabText.CodeBoxShapeNameRegex);
                if (shapes.Count > 0)
                {
                    codeShape = shapes[0];
                }
            }
            if (isURL)
            {
                codeBoxItem = new CodeBox(id,
                    codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], "", "", isURL, false, false, slide);
            }
            else if (isFile)
            {
                codeBoxItem = new CodeBox(id,
                    "", codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], "", false, isFile, false, slide);
            }
            else
            {
                codeBoxItem = new CodeBox(id,
                    "", "", codeBoxItemDic[LiveCodingLabText.CodeTextIdentifier], false, false, isText, slide);
            }

            if (codeShape != null)
            {
                codeBoxItem.Shape = codeShape;
            }

            CodeBoxPaneItem codeBoxPaneItem = new CodeBoxPaneItem(this, codeBoxItem);
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
        private void ClickHandler(Action<PowerPoint.ShapeRange> liveCodingAction, int minNoOfSelectedShapes, string[] errorParameters)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes == null || selectedShapes.Count < minNoOfSelectedShapes)
            {
                _errorHandler.ProcessErrorCode(LiveCodingLabErrorHandler.ErrorCodeInvalidSelection, errorParameters);
                return;
            }

            ExecuteLiveCodingAction(selectedShapes, liveCodingAction);
        }

        #endregion
    }
}
