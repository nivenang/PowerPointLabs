using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.LiveCodingLab;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.TextCollection;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class LiveCodingLabController : MarshalByRefObject, ILiveCodingLabController
    {
        private static ILiveCodingLabController _instance = new LiveCodingLabController();

        public static ILiveCodingLabController Instance { get { return _instance; } }

        private LiveCodingLabPane _pane;

        private LiveCodingLabController() { }

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(LiveCodingLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(LiveCodingLabPane)).Control as LiveCodingLabPane;
            }));
        }

        public void InsertTextCodeBox(string text)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    CodeBoxPaneItem item1 = CreateCodeBox();
                    item1.codeTextBox.Text = text;
                    item1.Text = text;
                    _pane.LiveCodingLabPaneWPF.RefreshCode();
                    InsertCodeBoxIntoSlide(item1);
                    item1.CodeBox.Shape.Name = "InsertTextCodeBoxTestShape";
                    item1.CodeBox.ShapeName = "InsertTextCodeBoxTestShape";
                });
            }
        }

        public void InsertFileCodeBox(string filePath)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    CodeBoxPaneItem item1 = CreateCodeBox();
                    item1.isFile.RaiseEvent(new RoutedEventArgs(ToggleButton.CheckedEvent));
                    item1.codeTextBox.Text = filePath;
                    item1.Text = filePath;
                    _pane.LiveCodingLabPaneWPF.RefreshCode();
                    InsertCodeBoxIntoSlide(item1);
                    item1.CodeBox.Shape.Name = "InsertFileCodeBoxTestShape";
                    item1.CodeBox.ShapeName = "InsertFileCodeBoxTestShape";
                });
            }
        }
        
        public void InsertDiffCodeBox(string diffFilePath)
        {
            Action<string, LiveCodingPaneWPF, string> liveCodingAction = (diffPath, liveCodingPane, diffGroupName) => _pane.LiveCodingLabPaneWPF.GetLiveCodingLabMain().InsertDiff(diffPath, liveCodingPane, diffGroupName);
            _pane.LiveCodingLabPaneWPF.Dispatcher.Invoke(liveCodingAction, diffFilePath, _pane.LiveCodingLabPaneWPF, "InsertDiffCodeBoxTestGroup");
        }
        
        public void InsertCodeBoxIntoSlide(CodeBoxPaneItem item)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    item.insertButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                });
            }
        }

        private CodeBoxPaneItem CreateCodeBox()
        {
            _pane.LiveCodingLabPaneWPF.insertCodeBox.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            return _pane.LiveCodingLabPaneWPF.GetCodeBoxList().First();
        }
    }
}
