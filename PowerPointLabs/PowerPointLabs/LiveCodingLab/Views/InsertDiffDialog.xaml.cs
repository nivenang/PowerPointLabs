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
    /// Interaction logic for InsertDiffDialog.xaml
    /// </summary>
    public partial class InsertDiffDialog
    {
        private string diffFilePath;
        private string diffGroupName;
        public InsertDiffDialog()
        {
            InitializeComponent();
            diffFilePath = "";
            diffGroupName = "";
        }
        public string DiffFile
        {
            get { return diffFilePath; }
            set { diffFilePath = value; }
        }

        public string DiffGroup
        {
            get { return diffGroupName; }
            set { diffGroupName = value; }
        }
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            diffFilePath = diffFileInput.Text;
            diffGroupName = diffGroupNameInput.Text;
            DialogResult = true;
            Close();
        }
    }
}
