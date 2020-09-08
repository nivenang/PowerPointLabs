using System;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            _errorHandler = LiveCodingLabErrorHandler.InitializeErrorHandler(this);
            Focusable = true;
        }

        internal void InitialiseLogic()
        {
            if (_liveCodingLab == null)
            {
                _liveCodingLab = new LiveCodingLabMain();
            }
        }

        #region XAML-Binded Event Handler

        private void InsertCodeBoxButton_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void RefreshCodeButton_Click(object sender, RoutedEventArgs e)
        {

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

        }
        #endregion

        #region Helper Methods


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
