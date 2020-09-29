using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace PowerPointLabs.LiveCodingLab.Views
{
    /// <summary>
    /// Interaction logic for AnimationSettingsDialog.xaml
    /// </summary>
    public partial class AnimationSettingsDialog
    {

        public delegate void DialogConfirmedDelegate(Drawing.Color highlightColor,
                                                    Drawing.Color defaultColor,
                                                    float scrollSpeed);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private float lastScrollSpeed;

        public AnimationSettingsDialog()
        {
            InitializeComponent();
        }

        public AnimationSettingsDialog(Drawing.Color defaultHighlightColor,
                                            Drawing.Color defaultTextColor,
                                            float defaultScrollSpeed)
            : this()
        {
            textHighlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultHighlightColor));
            textDefaultColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultTextColor));
            scrollSpeedInput.Text = defaultScrollSpeed.ToString("f");
            scrollSpeedInput.ToolTip = LiveCodingLabText.SettingsScrollSpeedTooltip;
            scrollSpeedInput.SelectAll();

            lastScrollSpeed = defaultScrollSpeed;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Drawing.Color textHighlightColor = GraphicsUtil.DrawingColorFromBrush(textHighlightColorRect.Fill);
            Drawing.Color textDefaultColor = GraphicsUtil.DrawingColorFromBrush(textDefaultColorRect.Fill);
            ValidateScrollSpeed();
            DialogConfirmedHandler(textHighlightColor, textDefaultColor, float.Parse(scrollSpeedInput.Text));
            Close();
        }

        private void TextHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textHighlightColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void TextDefaultColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textDefaultColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textDefaultColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void ScrollSpeedInput_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateScrollSpeed();
        }

        private void ValidateScrollSpeed ()
        {
            float enteredValue;
            if (float.TryParse(scrollSpeedInput.Text, out enteredValue))
            {
                if (enteredValue < 0.01)
                {
                    enteredValue = 0.01f;
                }
                else if (enteredValue > 59.0)
                {
                    enteredValue = 59.0f;
                }
            }
            else
            {
                enteredValue = lastScrollSpeed;
            }
            scrollSpeedInput.Text = enteredValue.ToString("f");
            lastScrollSpeed = enteredValue;
        }
    }
}
