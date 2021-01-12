using System.Collections.Generic;
using System.Drawing.Text;
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
    /// Interaction logic for LiveCodingLabSettingsDialogBox.xaml
    /// </summary>
    public partial class CodeBoxSettingsDialog
    {

        public delegate void DialogConfirmedDelegate(Drawing.Color textColor,
                                                    string fontType,
                                                    string fontSize,
                                                    string defaultLanguage);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private int lastFontSize;

        #region Constructors
        public CodeBoxSettingsDialog()
        {
            InitializeComponent();
            List<string> fonts = new List<string>();
            List<string> languages = new List<string>
            { 
                "Java", "Python", "C", "C++", "None"
            };

            foreach (Drawing.FontFamily font in Drawing.FontFamily.Families)
            {
                fonts.Add(font.Name);
            }
            fontComboBox.ItemsSource = fonts;
            languageComboBox.ItemsSource = languages;
        }

        public CodeBoxSettingsDialog(Drawing.Color defaultTextColor,
                                            string defaultFontType,
                                            string defaultFontSize,
                                            string defaultLanguage)
            : this()
        {
            textColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultTextColor));
            fontComboBox.Text = defaultFontType;
            fontSizeInput.Text = defaultFontSize;
            languageComboBox.Text = defaultLanguage;
        }
        #endregion

        #region Event Handlers
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Drawing.Color textColor = GraphicsUtil.DrawingColorFromBrush(textColorRect.Fill);
            string fontType = fontComboBox.Text;
            string fontSize = fontSizeInput.Text;
            string language = languageComboBox.Text;
            ValidateFontSize();
            DialogConfirmedHandler(textColor, fontType, fontSize, language);
            Close();
        }

        private void TextColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void FontSizeInput_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateFontSize();
        }
        #endregion

        #region Helper Methods
        private void ValidateFontSize()
        {
            int enteredValue;
            if (int.TryParse(fontSizeInput.Text, out enteredValue))
            {
                if (enteredValue < 8)
                {
                    enteredValue = 8;
                }
                else if (enteredValue > 96)
                {
                    enteredValue = 96;
                }
            }
            else
            {
                enteredValue = lastFontSize;
            }
            fontSizeInput.Text = enteredValue.ToString();
            lastFontSize = enteredValue;
        }
        #endregion
    }
}
