using System;
using System.Collections.Generic;

using PowerPointLabs.LiveCodingLab.Views;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{

    public interface ILiveCodingLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
        void Reset();
        void ExecuteLiveCodingAction(List<CodeBoxPaneItem> listCodeBox, Action<List<CodeBoxPaneItem>> liveCodingAction);
    }
}