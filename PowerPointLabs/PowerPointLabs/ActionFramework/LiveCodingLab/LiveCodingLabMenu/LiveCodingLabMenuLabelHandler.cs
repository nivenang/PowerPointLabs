using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Label.LiveCodingLab
{
    [ExportLabelRibbonId(LiveCodingLabText.RibbonMenuId)]
    class LiveCodingLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return LiveCodingLabText.RibbonMenuLabel;
        }
    }
}
