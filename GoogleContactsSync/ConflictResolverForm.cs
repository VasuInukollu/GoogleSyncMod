using System;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    internal partial class ConflictResolverForm : Form
    {
        public ConflictResolverForm()
        {
            InitializeComponent();
        }

        private void GoogleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GoogleComboBox.SelectedItem != null)
                GoogleItemTextBox.Text = ContactMatch.GetSummary((Google.Contacts.Contact)GoogleComboBox.SelectedItem);
        }

        private void ConflictResolverForm_Shown(object sender, EventArgs e)
        {
            SettingsForm.Instance.ShowBalloonToolTip(Text, messageLabel.Text, ToolTipIcon.Warning, 5000, true);

        }
    }
}