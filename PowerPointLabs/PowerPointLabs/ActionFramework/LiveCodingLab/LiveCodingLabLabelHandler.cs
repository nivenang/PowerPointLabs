using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportLabelRibbonId(LiveCodingLabText.LiveCodingLabPaneTag)]
    class LiveCodingLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return LiveCodingLabText.RibbonMenuLabel;
        }
    }
}
