using System;
using System.Collections.Generic;
using Google.Contacts;
using Google.Apis.Calendar.v3.Data;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    class ConflictResolver : IConflictResolver, IDisposable
    {
        private ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }

        #region IConflictResolver Members

        public ConflictResolution Resolve(ContactMatch match, bool isNewMatch)
        {
            string name = match.ToString();

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these Outlook and Google Contacts \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                    "Both the Outlook Contact and the Google Contact \"" + name +
                    "\" have been changed. Choose which you would like to keep.";
            }

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (match.OutlookContact != null)
            {
                Outlook.ContactItem item = match.OutlookContact.GetOriginalItemFromOutlook();
                try
                {
                    _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
                }
                finally
                {
                    if (item != null)
                    {
                        Marshal.ReleaseComObject(item);
                        item = null;
                    }
                }
            }

            if (match.GoogleContact != null)
                _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(match.GoogleContact);

            return Resolve();
        }

        public ConflictResolution ResolveDuplicate(OutlookContactInfo outlookContact, List<Contact> googleContacts, out Contact googleContact)
        {
            string name = ContactMatch.GetName(outlookContact);

            _form.messageLabel.Text =
                     "There are multiple Google Contacts (" + googleContacts.Count + ") matching unique properties for Outlook Contact \"" + name +
                     "\". Please choose from the combobox below the Google Contact you would like to match with Outlook and if you want to keep the Google or Outlook properties of the selected contact.";


            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;

            Outlook.ContactItem item = outlookContact.GetOriginalItemFromOutlook();
            try
            {
                _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            }
            finally
            {
                if (item != null)
                {
                    Marshal.ReleaseComObject(item);
                    item = null;
                }
            }

            _form.GoogleComboBox.DataSource = googleContacts;
            _form.GoogleComboBox.Visible = true;
            _form.AllCheckBox.Visible = false;
            _form.skip.Text = "Keep both";

            ConflictResolution res = Resolve();
            googleContact = _form.GoogleComboBox.SelectedItem as Contact;

            return res;
        }

        public DeleteResolution ResolveDelete(OutlookContactInfo outlookContact)
        {
            string name = ContactMatch.GetName(outlookContact);

            _form.Text = "Google Contact deleted";
            _form.messageLabel.Text =
                "Google Contact \"" + name +
                "\" doesn't exist anymore. Do you want to delete it also on Outlook side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            Outlook.ContactItem item = outlookContact.GetOriginalItemFromOutlook();
            try
            {
                _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            }
            finally
            {
                if (item != null)
                {
                    Marshal.ReleaseComObject(item);
                    item = null;
                }
            }

            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Contact googleContact)
        {
            string name = ContactMatch.GetName(googleContact);

            _form.Text = "Outlook Contact deleted";
            _form.messageLabel.Text =
                "Outlook Contact \"" + name +
                "\" doesn't exist anymore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(googleContact);

            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        private ConflictResolution Resolve()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.Ignore:
                    // skip
                    return _form.AllCheckBox.Checked ? ConflictResolution.SkipAlways : ConflictResolution.Skip;
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.GoogleWinsAlways : ConflictResolution.GoogleWins;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.OutlookWinsAlways : ConflictResolution.OutlookWins;
                default:
                    return ConflictResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedOutlook()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteGoogleAlways : DeleteResolution.DeleteGoogle;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepGoogleAlways : DeleteResolution.KeepGoogle;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedGoogle()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteOutlookAlways : DeleteResolution.DeleteOutlook;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepOutlookAlways : DeleteResolution.KeepOutlook;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        public ConflictResolution Resolve(Outlook.AppointmentItem outlookAppointment, Event googleAppointment, AppointmentsSynchronizer sync, bool isNewMatch)
        {
            string name = string.Empty;

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (outlookAppointment != null)
            {
                name = outlookAppointment.Subject + " - " + outlookAppointment.Start;
                _form.OutlookItemTextBox.Text += outlookAppointment.Body;
            }

            if (googleAppointment != null)
            {
                name = googleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(googleAppointment);
                _form.GoogleItemTextBox.Text += googleAppointment.Description;
            }

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these appointments \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                "Both the Outlook and Google Appointment \"" + name +
                "\" have been changed. Choose which you would like to keep.";
            }

            return Resolve();
        }

        public ConflictResolution Resolve(string message, Outlook.AppointmentItem outlookAppointment, Event googleAppointment, AppointmentsSynchronizer sync, bool keepOutlook, bool keepGoogle)
        {
            // string name = string.Empty;

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (outlookAppointment != null)
            {
                // name = outlookAppointment.Subject + " - " + outlookAppointment.Start;
                _form.OutlookItemTextBox.Text += outlookAppointment.Body;
            }

            if (googleAppointment != null)
            {
                // name = googleAppointment.Summary + " - " + Synchronizer.GetTime(googleAppointment);
                _form.GoogleItemTextBox.Text += googleAppointment.Description;
            }

            //ToDo: Make it more flexible
            _form.keepGoogle.Enabled = keepGoogle;
            _form.keepOutlook.Enabled = keepOutlook;
            _form.AllCheckBox.Visible = true;
            _form.messageLabel.Text = message;

            return Resolve();
        }

        public ConflictResolution Resolve(string message, Outlook.AppointmentItem outlookAppointment, Event googleAppointment, AppointmentsSynchronizer sync)
        {
            return Resolve(message, outlookAppointment, googleAppointment, sync, true, false);
        }

        public ConflictResolution Resolve(string message, Event googleAppointment, Outlook.AppointmentItem outlookAppointment, AppointmentsSynchronizer sync)
        {
            return Resolve(message, outlookAppointment, googleAppointment, sync, false, true);
        }

        public DeleteResolution ResolveDelete(Outlook.AppointmentItem outlookAppointment)
        {
            _form.Text = "Google appointment deleted";
            _form.messageLabel.Text =
                "Google appointment \"" + outlookAppointment.Subject + " - " + outlookAppointment.Start +
                "\" doesn't exist anymore. Do you want to delete it also on Outlook side?";

            _form.GoogleItemTextBox.Text = string.Empty;
            _form.OutlookItemTextBox.Text += outlookAppointment.Body;

            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Event googleAppointment)
        {
            _form.Text = "Outlook appointment deleted";
            _form.messageLabel.Text =
                "Outlook appointment \"" + googleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(googleAppointment) +
                "\" doesn't exist anymore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text += googleAppointment.Description;


            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        public void Dispose()
        {
            ((IDisposable)_form).Dispose();
        }

        #endregion
    }
}
