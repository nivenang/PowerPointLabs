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
    /// Interaction logic for GroupCodeBoxDialog.xaml
    /// </summary>
    public partial class GroupCodeBoxDialog
    {
        private string groupName;
        public GroupCodeBoxDialog()
        {
            InitializeComponent();
            groupName = "";
        }
        public string ResponseText
        {
            get { return groupName; }
            set { groupName = value; }
        }
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            groupName = groupNameInput.Text;
            DialogResult = true;
            Close();
        }
    }
}
