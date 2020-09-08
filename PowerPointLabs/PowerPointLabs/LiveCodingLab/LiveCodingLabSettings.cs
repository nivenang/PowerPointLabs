using System.Drawing;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.LiveCodingLab.Views;

namespace PowerPointLabs.LiveCodingLab
{
    internal static class LiveCodingLabSettings
    {

        public static Color bulletsTextHighlightColor = Color.FromArgb(242, 41, 10);
        public static Color bulletsTextDefaultColor = Color.FromArgb(0, 0, 0);
        public static float scrollSpeedDefaultValue = 1.0f;

        public static void ShowAnimationSettingsDialog()
        {
            AnimationSettingsDialog dialog = new AnimationSettingsDialog(
                bulletsTextHighlightColor, 
                bulletsTextDefaultColor, 
                scrollSpeedDefaultValue);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnSettingsDialogConfirmed(Color newHighlightColor, Color newDefaultColor, float newScrollSpeed)
        {
            bulletsTextHighlightColor = newHighlightColor;
            bulletsTextDefaultColor = newDefaultColor;
            scrollSpeedDefaultValue = newScrollSpeed;
        }
    }
}
