using System;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Drawing;

namespace GoContactSyncMod
{
    public partial class ConfigurationManagerForm : Form
    {
        public ConfigurationManagerForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
        }

        public string AddProfile()
        {
            string vReturn = "";
            using (AddEditProfileForm AddEditProfile = new AddEditProfileForm("New profile", null))
            {
                if (AddEditProfile.ShowDialog(SettingsForm.Instance) == DialogResult.OK)
                {
                    if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                    {
                        MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "New profile");
                    }
                    else
                    {
                        Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName);
                        vReturn = AddEditProfile.ProfileName;
                    }
                }
            }
            return vReturn;
        }

        private void fillListProfiles()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey);

            lbProfiles.Items.Clear();

            foreach (string subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                lbProfiles.Items.Add(subKeyName);
            }
        }

        //copy all the values
        private static void CopyKey(RegistryKey parent, string keyNameSource, string keyNameDestination)
        {
            RegistryKey destination = parent.CreateSubKey(keyNameDestination);
            RegistryKey source = parent.OpenSubKey(keyNameSource);

            foreach (string valueName in source.GetValueNames())
            {
                object objValue = source.GetValue(valueName);
                RegistryValueKind valKind = source.GetValueKind(valueName);
                destination.SetValue(valueName, objValue, valKind);
            }
        }

        private void btClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            AddProfile();
            fillListProfiles();
        }

        private void btEdit_Click(object sender, EventArgs e)
        {
            if (1 == lbProfiles.CheckedItems.Count)
            {
                using (AddEditProfileForm AddEditProfile = new AddEditProfileForm("Edit profile", lbProfiles.CheckedItems[0].ToString()))
                {
                    if (AddEditProfile.ShowDialog(SettingsForm.Instance) == DialogResult.OK)
                    {
                        if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                        {
                            MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "Edit profile");
                        }
                        else
                        {
                            CopyKey(Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey), lbProfiles.CheckedItems[0].ToString(), AddEditProfile.ProfileName);
                            Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + lbProfiles.CheckedItems[0].ToString());
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, select one profile for editing", "Edit profile");
            }

            fillListProfiles();
        }

        private void btDel_Click(object sender, EventArgs e)
        {
            if (0 >= lbProfiles.CheckedItems.Count)
            {
                MessageBox.Show("You don`t select any profile. Deletion imposble.", "Delete profile");
            }
            else if (DialogResult.Yes == MessageBox.Show("Do you sure to delete selection ?", "Delete profile",
                                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                foreach (object itemChecked in lbProfiles.CheckedItems)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + itemChecked.ToString());
                }
            }

            fillListProfiles();
        }

        private void ConfigurationManagerForm_Load(object sender, EventArgs e)
        {
            fillListProfiles();
        }
    }
}
