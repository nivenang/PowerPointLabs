using System.Drawing;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.LiveCodingLab.Views;

namespace PowerPointLabs.LiveCodingLab
{
    internal static class LiveCodingLabSettings
    {

        public static Color bulletsTextHighlightColor = Color.FromArgb(242, 41, 10);
        public static Color bulletsTextDefaultColor = Color.FromArgb(0, 0, 0);
        public static Color codeTextColor = Color.FromArgb(0, 0, 0);
        public static string codeFontType = "Consolas";
        public static string language = "None";
        public static int codeFontSize = 18;

        public static void ShowCodeBoxSettingsDialog()
        {
            CodeBoxSettingsDialog dialog = new CodeBoxSettingsDialog(
                codeTextColor,
                codeFontType,
                codeFontSize.ToString(),
                language);
            dialog.DialogConfirmedHandler += OnCodeBoxSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }
        public static void ShowAnimationSettingsDialog()
        {
            AnimationSettingsDialog dialog = new AnimationSettingsDialog(
                bulletsTextHighlightColor,
                bulletsTextDefaultColor);
            dialog.DialogConfirmedHandler += OnAnimationSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnCodeBoxSettingsDialogConfirmed(Color newCodeColor, string newFontType, string newFontSize, string newLanguage)
        {
            codeTextColor = newCodeColor;
            codeFontType = newFontType;
            codeFontSize = int.Parse(newFontSize);
            language = newLanguage;
        }

        private static void OnAnimationSettingsDialogConfirmed(Color newHighlightColor, Color newDefaultColor)
        {
            bulletsTextHighlightColor = newHighlightColor;
            bulletsTextDefaultColor = newDefaultColor;
        }
    }
}
