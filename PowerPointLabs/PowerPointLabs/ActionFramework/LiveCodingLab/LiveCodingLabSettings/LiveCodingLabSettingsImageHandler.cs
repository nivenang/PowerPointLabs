using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Image.LiveCodingLab
{
    [ExportImageRibbonId(LiveCodingLabText.SettingsTag)]
    class LiveCodingLabSettingsImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.SettingsGearBlue);
        }
    }
}
