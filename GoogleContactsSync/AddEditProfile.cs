using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class AddEditProfileForm : Form
    {
        public string ProfileName
        {
            get { return tbProfileName.Text; }
        }

        public AddEditProfileForm()
        {
            InitializeComponent();
        }

        public AddEditProfileForm(string title, string profileName)
        {
            InitializeComponent();

            if (!string.IsNullOrEmpty(title))
                Text = title;

            if (!string.IsNullOrEmpty(profileName))
                tbProfileName.Text = profileName;
        }

    }
}
