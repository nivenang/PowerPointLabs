﻿using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Image.LiveCodingLab
{
    [ExportImageRibbonId(LiveCodingLabText.RibbonMenuId)]
    class LiveCodingLabMenuImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.LiveCodingLab);
        }
    }
}
