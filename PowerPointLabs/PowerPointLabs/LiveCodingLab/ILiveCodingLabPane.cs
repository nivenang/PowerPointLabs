using System;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{

    public interface ILiveCodingLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
        void Reset();
        void ExecuteLiveCodingAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> liveCodingAction);
    }
}