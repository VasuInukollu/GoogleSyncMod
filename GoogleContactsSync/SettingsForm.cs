using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using Google.Apis.Util.Store;
using Google.Apis.Calendar.v3.Data;
using System.Threading.Tasks;

namespace GoContactSyncMod
{
    internal partial class SettingsForm : Form
    {
        //Singleton-Object
        #region Singleton Definition

        private static volatile SettingsForm instance;
        private static object syncRoot = new Object();

        public static SettingsForm Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new SettingsForm();
                    }
                }
                return instance;
            }
        }
        #endregion

        internal Synchronizer sync;
        private SyncOption syncOption;
        private DateTime lastSync;
        private bool requestClose = false;
        private bool boolShowBalloonTip = true;
        private CancellationTokenSource cancellationTokenSource;

        public const string AppRootKey = @"Software\GoContactSyncMOD";
        public const string RegistrySyncOption = "SyncOption";
        public const string RegistryUsername = "Username";
        public const string RegistryAutoSync = "AutoSync";
        public const string RegistryAutoSyncInterval = "AutoSyncInterval";
        public const string RegistryAutoStart = "AutoStart";
        public const string RegistryReportSyncResult = "ReportSyncResult";
        public const string RegistrySyncDeletion = "SyncDeletion";
        public const string RegistryPromptDeletion = "PromptDeletion";
        public const string RegistrySyncAppointmentsMonthsInPast = "SyncAppointmentsMonthsInPast";
        public const string RegistrySyncAppointmentsMonthsInFuture = "SyncAppointmentsMonthsInFuture";
        public const string RegistrySyncAppointmentsTimezone = "SyncAppointmentsTimezone";
        public const string RegistrySyncAppointments = "SyncAppointments";
        public const string RegistrySyncAppointmentsForceRTF = "SyncAppointmentsForceRTF";
        public const string RegistrySyncNotes = "SyncNotes";
        public const string RegistrySyncContacts = "SyncContacts";
        public const string RegistrySyncContactsForceRTF = "SyncContactsForceRTF";
        public const string RegistryUseFileAs = "UseFileAs";
        public const string RegistryLastSync = "LastSync";
        public const string RegistrySyncContactsFolder = "SyncContactsFolder";
        public const string RegistrySyncNotesFolder = "SyncNotesFolder";
        public const string RegistrySyncAppointmentsFolder = "SyncAppointmentsFolder";
        public const string RegistrySyncAppointmentsGoogleFolder = "SyncAppointmentsGoogleFolder";
        public const string RegistrySyncProfile = "SyncProfile";

        private ProxySettingsForm _proxy = new ProxySettingsForm();

        private string syncContactsFolder = "";
        private string syncNotesFolder = "";
        private string syncAppointmentsFolder = "";
        private string syncAppointmentsGoogleFolder = "";
        private string Timezone = "";

        //private string _syncProfile;
        private static string SyncProfile
        {
            get
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                return (regKeyAppRoot.GetValue(RegistrySyncProfile) != null) ?
                       (string)regKeyAppRoot.GetValue(RegistrySyncProfile) : null;
            }
            set
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                if (value != null)
                {
                    regKeyAppRoot.SetValue(RegistrySyncProfile, value);
                }
            }
        }


        private int executing; // make this static if you want this one-caller-only to
        // all objects instead of a single object

        Thread syncThread;

        //register window for lock/unlock messages of workstation
        //private bool registered = false;

        delegate void TextHandler(string text);
        delegate void SwitchHandler(bool value);
        delegate void IconHandler();
        delegate DialogResult DialogHandler(string text);
        delegate void OnTimeZoneChangesCallback(string timeZone);

        public DialogResult ShowDialog(string text)
        {
            if (this.InvokeRequired)
            {
                return (DialogResult)Invoke(new DialogHandler(ShowDialog), new object[] { text });
            }
            else
            {
                return MessageBox.Show(this, text, Application.ProductName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            }


        }

        private Icon IconError = GoContactSyncMod.Properties.Resources.sync_error;
        private Icon Icon0 = GoContactSyncMod.Properties.Resources.sync;
        private Icon Icon30 = GoContactSyncMod.Properties.Resources.sync_30;
        private Icon Icon60 = GoContactSyncMod.Properties.Resources.sync_60;
        private Icon Icon90 = GoContactSyncMod.Properties.Resources.sync_90;
        private Icon Icon120 = GoContactSyncMod.Properties.Resources.sync_120;
        private Icon Icon150 = GoContactSyncMod.Properties.Resources.sync_150;
        private Icon Icon180 = GoContactSyncMod.Properties.Resources.sync_180;
        private Icon Icon210 = GoContactSyncMod.Properties.Resources.sync_210;
        private Icon Icon240 = GoContactSyncMod.Properties.Resources.sync_240;
        private Icon Icon270 = GoContactSyncMod.Properties.Resources.sync_270;
        private Icon Icon300 = GoContactSyncMod.Properties.Resources.sync_300;
        private Icon Icon330 = GoContactSyncMod.Properties.Resources.sync_330;

        private SettingsForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            cancellationTokenSource = new CancellationTokenSource();
            InitializeComponent();
            Text = Text + " - " + Application.ProductVersion;
            Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
            Logger.Log("Started application " + Application.ProductName + " (" + Application.ProductVersion + ") on " + VersionInformation.GetWindowsVersion() + " and " + OutlookRegistryUtils.GetOutlookVersion(), EventType.Information);
            Logger.Log("Detailed log file created: " + Logger.Folder + "log.txt", EventType.Information);
            ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            NotesMatcher.NotificationReceived += new NotesMatcher.NotificationHandler(OnNotificationReceived);
            AppointmentsMatcher.NotificationReceived += new AppointmentsMatcher.NotificationHandler(OnNotificationReceived);
            PopulateSyncOptionBox();

            //temporary remove the listener to avoid to load the settings twice, because it is set from SettingsForm.Designer.cs
            this.cmbSyncProfile.SelectedIndexChanged -= new System.EventHandler(this.cmbSyncProfile_SelectedIndexChanged);
            if (fillSyncProfileItems())
                LoadSettings(cmbSyncProfile.Text);
            else
                LoadSettings(null);
            //enable the listener
            this.cmbSyncProfile.SelectedIndexChanged += new System.EventHandler(this.cmbSyncProfile_SelectedIndexChanged);

            TimerSwitch(true);
            lastSyncLabel.Text = "Not synced";

            ValidateSyncButton();

            //Register Session Lock Event
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            //Register Power Mode Event
            SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler(SystemEvents_PowerModeSwitch);
        }

        private void PopulateSyncOptionBox()
        {
            string str;
            for (int i = 0; i < 20; i++)
            {
                str = ((SyncOption)i).ToString();
                if (str == i.ToString())
                    break;

                // format (to add space before capital)
                MatchCollection matches = Regex.Matches(str, "[A-Z]");
                for (int k = 0; k < matches.Count; k++)
                {
                    str = str.Replace(str[matches[k].Index].ToString(), " " + str[matches[k].Index]);
                    matches = Regex.Matches(str, "[A-Z]");
                }
                str = str.Replace("  ", " ");
                // fix start
                str = str.Substring(1);

                syncOptionBox.Items.Add(str);
            }
        }
        private void fillSyncFolderItems()
        {
            if (this.InvokeRequired)
            {
                Invoke(new InvokeCallback(fillSyncFolderItems));
            }
            else
            {

                lock (syncRoot)
                {
                    if (this.contactFoldersComboBox.DataSource == null || /*this.noteFoldersComboBox.DataSource == null ||*/ this.appointmentFoldersComboBox.DataSource == null || this.appointmentGoogleFoldersComboBox.DataSource == null && btSyncAppointments.Checked ||
                        this.contactFoldersComboBox.Items.Count == 0 || /*this.noteFoldersComboBox.Items.Count == 0 ||*/ this.appointmentFoldersComboBox.Items.Count == 0 || this.appointmentGoogleFoldersComboBox.Items.Count == 0 && btSyncAppointments.Checked)
                    {//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                        Logger.Log("Loading Outlook folders...", EventType.Information);

                        this.contactFoldersComboBox.Visible = this.btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
                        //this.noteFoldersComboBox.Visible = btSyncNotes.Checked;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                        this.labelTimezone.Visible = this.labelMonthsPast.Visible = this.labelMonthsFuture.Visible = btSyncAppointments.Checked;
                        this.appointmentFoldersComboBox.Visible = this.appointmentGoogleFoldersComboBox.Visible = this.futureMonthInterval.Visible = this.pastMonthInterval.Visible = this.appointmentTimezonesComboBox.Visible = this.btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
                        this.cmbSyncProfile.Visible = true;

                        string defaultText = "    --- Select an Outlook folder ---";
                        ArrayList outlookContactFolders = new ArrayList();
                        ArrayList outlookNoteFolders = new ArrayList();
                        ArrayList outlookAppointmentFolders = new ArrayList();

                        try
                        {
                            Cursor = Cursors.WaitCursor;
                            SuspendLayout();

                            this.contactFoldersComboBox.BeginUpdate();
                            //this.noteFoldersComboBox.BeginUpdate();//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            this.appointmentFoldersComboBox.BeginUpdate();
                            this.contactFoldersComboBox.DataSource = null;
                            //this.noteFoldersComboBox.DataSource = null;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            this.appointmentFoldersComboBox.DataSource = null;
                            //this.contactFoldersComboBox.Items.Clear();

                            Microsoft.Office.Interop.Outlook.Folders folders = Synchronizer.OutlookNameSpace.Folders;
                            foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
                            {
                                try
                                {
                                    GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, outlookAppointmentFolders, folder);
                                }
                                catch (Exception e)
                                {
                                    Logger.Log(e, EventType.Debug);
                                    Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                                }
                            }

                            if (outlookContactFolders != null) // && outlookContactFolders.Count > 0)
                            {
                                outlookContactFolders.Sort();
                                outlookContactFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                this.contactFoldersComboBox.DataSource = outlookContactFolders;
                                this.contactFoldersComboBox.DisplayMember = "DisplayName";
                                this.contactFoldersComboBox.ValueMember = "FolderID";
                            }


                            //ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            //if (outlookNoteFolders != null) // && outlookNoteFolders.Count > 0)
                            //{
                            //    outlookNoteFolders.Sort();
                            //    outlookNoteFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                            //    this.noteFoldersComboBox.DataSource = outlookNoteFolders;
                            //    this.noteFoldersComboBox.DisplayMember = "DisplayName";
                            //    this.noteFoldersComboBox.ValueMember = "FolderID";
                            //}

                            if (outlookAppointmentFolders != null) // && outlookAppointmentFolders.Count > 0)
                            {
                                outlookAppointmentFolders.Sort();
                                outlookAppointmentFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                this.appointmentFoldersComboBox.DataSource = outlookAppointmentFolders;
                                this.appointmentFoldersComboBox.DisplayMember = "DisplayName";
                                this.appointmentFoldersComboBox.ValueMember = "FolderID";
                            }

                            this.contactFoldersComboBox.EndUpdate();
                            //this.noteFoldersComboBox.EndUpdate();//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            this.appointmentFoldersComboBox.EndUpdate();

                            this.contactFoldersComboBox.SelectedValue = defaultText;
                            //this.noteFoldersComboBox.SelectedValue = defaultText;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            this.appointmentFoldersComboBox.SelectedValue = defaultText;

                            //this.contactFoldersComboBox.SelectedValue = "";
                            //this.noteFoldersComboBox.SelectedValue = "";
                            //this.appointmentFoldersComboBox.SelectedValue = "";

                            //Select Default Folder per Default
                            foreach (OutlookFolder folder in contactFoldersComboBox.Items)
                                if (folder.IsDefaultFolder)
                                {
                                    this.contactFoldersComboBox.SelectedValue = folder.FolderID;
                                    break;
                                }

                            //Select Default Folder per Default
                            //ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                            //foreach (OutlookFolder folder in noteFoldersComboBox.Items)
                            //    if (folder.IsDefaultFolder)
                            //    {
                            //        this.noteFoldersComboBox.SelectedItem = folder;
                            //        break;
                            //    }

                            //Select Default Folder per Default
                            foreach (OutlookFolder folder in appointmentFoldersComboBox.Items)
                                if (folder.IsDefaultFolder)
                                {
                                    this.appointmentFoldersComboBox.SelectedItem = folder;
                                    break;
                                }

                            Logger.Log("Loaded Outlook folders.", EventType.Information);

                        }

                        catch (Exception e)
                        {
                            Logger.Log(e, EventType.Debug);
                            Logger.Log("Error getting available Outlook and Google folders: " + e.Message, EventType.Warning);

                        }
                        finally
                        {
                            Cursor = Cursors.Default;
                            ResumeLayout();
                        }

                        LoadSettingsFolders(SyncProfile);

                        if ((contactFoldersComboBox.SelectedIndex == -1) && (contactFoldersComboBox.Items.Count > 0))
                            contactFoldersComboBox.SelectedIndex = 0;

                        if ((appointmentFoldersComboBox.SelectedIndex == -1) && (appointmentFoldersComboBox.Items.Count > 0))
                            appointmentFoldersComboBox.SelectedIndex = 0;
                    }
                }
            }
        }

        public static void GetOutlookMAPIFolders(ArrayList outlookContactFolders, ArrayList outlookNoteFolders, ArrayList outlookAppointmentFolders, Microsoft.Office.Interop.Outlook.MAPIFolder folder)
        {
            foreach (Microsoft.Office.Interop.Outlook.MAPIFolder mapi in folder.Folders)
            {
                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                {
                    bool isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts).EntryID);
                    outlookContactFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                }
                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olNoteItem)
                {
                    bool isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderNotes).EntryID);
                    outlookNoteFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                }
                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                {
                    bool isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar).EntryID);
                    outlookAppointmentFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                }

                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem ||
                    mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olNoteItem ||
                    mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                    GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, outlookAppointmentFolders, mapi);

            }
        }

        private void ClearSettings()
        {
            SetSyncOption(0);
            autoSyncCheckBox.Checked = runAtStartupCheckBox.Checked = reportSyncResultCheckBox.Checked = false;
            autoSyncInterval.Value = 120;
            _proxy.ClearSettings();
        }
        // Fill lists of sync profiles
        private bool fillSyncProfileItems()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            }


            bool vReturn = false;

            cmbSyncProfile.Items.Clear();
            cmbSyncProfile.Items.Add("[Add new profile...]");

            foreach (string subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                if (!string.IsNullOrEmpty(subKeyName))
                    cmbSyncProfile.Items.Add(subKeyName);
            }

            if (SyncProfile == null)
                SyncProfile = "Default_" + System.Environment.MachineName;

            if (cmbSyncProfile.Items.Count == 1)
                cmbSyncProfile.Items.Add(SyncProfile);
            else
                vReturn = true;

            cmbSyncProfile.Items.Add("[Configuration manager...]");
            cmbSyncProfile.Text = SyncProfile;

            return vReturn;
        }


        private void LoadSettings(string _profile)
        {
            Logger.Log("Loading settings from registry...", EventType.Information);
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                //MessageBox.Show("Your settings have been deleted because of an upgrade! You simply need to reconfigure them. Thx!", Application.ProductName + " - INFORMATION",MessageBoxButtons.OK);
                //Registry.CurrentUser.DeleteSubKeyTree(@"Software\Webgear\GOContactSync");
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync" + (_profile != null ? ('\\' + _profile) : ""));
            }

            if (regKeyAppRoot.GetValue(RegistrySyncOption) != null)
            {
                syncOption = (SyncOption)regKeyAppRoot.GetValue(RegistrySyncOption);
                SetSyncOption((int)syncOption);
            }

            if (regKeyAppRoot.GetValue(RegistryUsername) != null)
            {
                UserName.Text = regKeyAppRoot.GetValue(RegistryUsername) as string;
                //if (regKeyAppRoot.GetValue("Password") != null)
                //    Password.Text = Encryption.DecryptPassword(UserName.Text, regKeyAppRoot.GetValue("Password") as string);
            }
            //if (regKeyAppRoot.GetValue("Password") != null)
            //{
            //    regKeyAppRoot.DeleteValue("Password");
            //}

            //temporary remove listener
            this.autoSyncCheckBox.CheckedChanged -= new System.EventHandler(this.autoSyncCheckBox_CheckedChanged);

            ReadRegistryIntoCheckBox(autoSyncCheckBox, regKeyAppRoot.GetValue(RegistryAutoSync));
            ReadRegistryIntoNumber(autoSyncInterval, regKeyAppRoot.GetValue(RegistryAutoSyncInterval));
            ReadRegistryIntoCheckBox(runAtStartupCheckBox, regKeyAppRoot.GetValue(RegistryAutoStart));
            ReadRegistryIntoCheckBox(reportSyncResultCheckBox, regKeyAppRoot.GetValue(RegistryReportSyncResult));
            ReadRegistryIntoCheckBox(btSyncDelete, regKeyAppRoot.GetValue(RegistrySyncDeletion));
            ReadRegistryIntoCheckBox(btPromptDelete, regKeyAppRoot.GetValue(RegistryPromptDeletion));
            ReadRegistryIntoNumber(pastMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInPast));
            ReadRegistryIntoNumber(futureMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInFuture));
            if (regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) != null)
                appointmentTimezonesComboBox.Text = regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) as string;
            ReadRegistryIntoCheckBox(btSyncAppointments, regKeyAppRoot.GetValue(RegistrySyncAppointments));
            //ReadRegistryIntoCheckBox(btSyncNotes, regKeyAppRoot.GetValue(RegistrySyncNotes));//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
            object registryEntry = regKeyAppRoot.GetValue(RegistrySyncNotes);
            if (registryEntry != null)
            {
                try
                {
                    bool syncNotes = Convert.ToBoolean(registryEntry);
                    if (syncNotes)
                        Logger.Log("Notes Sync doesn't work anymore, because Google.Documents API was replaced by Google.Drive API on 21-Apr-2015 and it is not compatible. Thefore Notes Sync was removed from GCSM.", EventType.Information);
                }
                catch (Exception)
                {
                    //ignored;
                }
            }

            ReadRegistryIntoCheckBox(btSyncContacts, regKeyAppRoot.GetValue(RegistrySyncContacts));
            ReadRegistryIntoCheckBox(chkUseFileAs, regKeyAppRoot.GetValue(RegistryUseFileAs));

            ReadRegistryIntoCheckBox(btSyncContactsForceRTF, regKeyAppRoot.GetValue(RegistrySyncContactsForceRTF));
            ReadRegistryIntoCheckBox(btSyncAppointmentsForceRTF, regKeyAppRoot.GetValue(RegistrySyncAppointmentsForceRTF));

            if (regKeyAppRoot.GetValue(RegistryLastSync) != null)
            {
                try
                {
                    lastSync = new DateTime(Convert.ToInt64(regKeyAppRoot.GetValue(RegistryLastSync)));
                    SetLastSyncText(lastSync.ToString());
                }
                catch (System.FormatException ex)
                {
                    Logger.Log("LastSyncDate couldn't be read from registry (" + regKeyAppRoot.GetValue(RegistryLastSync) + "): " + ex, EventType.Warning);
                }
            }
            LoadSettingsFolders(_profile);

            //autoSyncCheckBox_CheckedChanged(null, null);
            btSyncContacts_CheckedChanged(null, null);
            //btSyncNotes_CheckedChanged(null, null);//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015

            _proxy.LoadSettings(_profile);

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                SaveSettings(_profile);
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Webgear\GOContactSync");
            }

            //enable temporary disabled listener
            this.autoSyncCheckBox.CheckedChanged += new System.EventHandler(this.autoSyncCheckBox_CheckedChanged);
        }

        private static void ReadRegistryIntoCheckBox(CheckBox checkbox, object registryEntry)
        {
            if (registryEntry != null)
            {
                try
                {
                    checkbox.Checked = Convert.ToBoolean(registryEntry);
                }
                catch (System.FormatException ex)
                {
                    Logger.Log(checkbox.Name + " couldn't be read from registry (" + registryEntry + "), was kept at default (" + checkbox.Checked + "): " + ex, EventType.Warning);

                }
            }

        }

        private static void ReadRegistryIntoNumber(NumericUpDown numericUpDown, object registryEntry)
        {
            if (registryEntry != null)
            {
                decimal interval = Convert.ToDecimal(registryEntry);
                if (interval < numericUpDown.Minimum)
                {
                    numericUpDown.Value = numericUpDown.Minimum;
                    Logger.Log(numericUpDown.Name + " read from registry was below range (" + interval + "), was set to minimum (" + numericUpDown.Minimum + ")", EventType.Warning);
                }
                else if (interval > numericUpDown.Maximum)
                {
                    numericUpDown.Value = numericUpDown.Maximum;
                    Logger.Log(numericUpDown.Name + " read from registry was above range (" + interval + "), was set to maximum (" + numericUpDown.Maximum + ")", EventType.Warning);
                }
                else
                    numericUpDown.Value = interval;
            }
        }

        private void LoadSettingsFolders(string _profile)
        {

            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync" + (_profile != null ? ('\\' + _profile) : ""));
            }

            string regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                if (contactFoldersComboBox.Items.Contains(regKeyValueStr))
                    contactFoldersComboBox.SelectedValue = regKeyValueStr;
            }
            //ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
            //regKeyValue = regKeyAppRoot.GetValue(RegistrySyncNotesFolder);
            //if (regKeyValue != null && !string.IsNullOrEmpty(regKeyValue as string))
            //    noteFoldersComboBox.SelectedValue = regKeyValue as string;

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                if (appointmentFoldersComboBox.Items.Contains(regKeyValueStr))
                    appointmentFoldersComboBox.SelectedValue = regKeyValueStr;
            }

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                if (appointmentGoogleFoldersComboBox.DataSource == null)
                {
                    appointmentFoldersComboBox.BeginUpdate();
                    ArrayList list = new ArrayList();
                    list.Add(new GoogleCalendar(regKeyValueStr, regKeyValueStr, false));
                    appointmentGoogleFoldersComboBox.DataSource = list;
                    appointmentGoogleFoldersComboBox.DisplayMember = "DisplayName";
                    appointmentGoogleFoldersComboBox.ValueMember = "FolderID";
                    //this.appointmentGoogleFoldersComboBox.SelectedIndex = 0;
                    appointmentFoldersComboBox.EndUpdate();
                }

                appointmentGoogleFoldersComboBox.SelectedValue = (regKeyValueStr);
            }
        }

        private void SaveSettings()
        {
            SaveSettings(cmbSyncProfile.Text);
        }

        private void SaveSettings(string profile)
        {
            if (!string.IsNullOrEmpty(profile))
            {
                SyncProfile = cmbSyncProfile.Text;
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + profile);
                regKeyAppRoot.SetValue(RegistrySyncOption, (int)syncOption);

                if (!string.IsNullOrEmpty(UserName.Text))
                {
                    regKeyAppRoot.SetValue(RegistryUsername, UserName.Text);
                }
                regKeyAppRoot.SetValue(RegistryAutoSync, autoSyncCheckBox.Checked.ToString());
                regKeyAppRoot.SetValue(RegistryAutoSyncInterval, autoSyncInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistryAutoStart, runAtStartupCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistryReportSyncResult, reportSyncResultCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistrySyncDeletion, btSyncDelete.Checked);
                regKeyAppRoot.SetValue(RegistryPromptDeletion, btPromptDelete.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInPast, pastMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInFuture, futureMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsTimezone, appointmentTimezonesComboBox.Text);
                regKeyAppRoot.SetValue(RegistrySyncAppointments, btSyncAppointments.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsForceRTF, btSyncAppointmentsForceRTF.Checked);
                // regKeyAppRoot.SetValue(RegistrySyncNotes, btSyncNotes.Checked);//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                regKeyAppRoot.SetValue(RegistrySyncContacts, btSyncContacts.Checked);
                regKeyAppRoot.SetValue(RegistrySyncContactsForceRTF, btSyncContactsForceRTF.Checked);
                regKeyAppRoot.SetValue(RegistryUseFileAs, chkUseFileAs.Checked);
                regKeyAppRoot.SetValue(RegistryLastSync, lastSync.Ticks);

                //if (btSyncContacts.Checked && contactFoldersComboBox.SelectedValue != null)
                //    regKeyAppRoot.SetValue("SyncContactsFolder", contactFoldersComboBox.SelectedValue.ToString());
                //if (btSyncNotes.Checked && noteFoldersComboBox.SelectedValue != null)
                //    regKeyAppRoot.SetValue("SyncNotesFolder", noteFoldersComboBox.SelectedValue.ToString());

                _proxy.SaveSettings(cmbSyncProfile.Text);
            }
        }

        private bool ValidSyncFolders
        {
            get
            {
                bool syncContactFolderIsValid = (contactFoldersComboBox.SelectedIndex >= 1 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count)
                                                || !btSyncContacts.Checked;
                //ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                //bool syncNoteFolderIsValid = (noteFoldersComboBox.SelectedIndex >= 1 && noteFoldersComboBox.SelectedIndex < noteFoldersComboBox.Items.Count)
                //                                || !btSyncNotes.Checked;
                bool syncAppointmentFolderIsValid = (appointmentFoldersComboBox.SelectedIndex >= 1 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count)
                        && (appointmentGoogleFoldersComboBox.SelectedIndex == appointmentGoogleFoldersComboBox.Items.Count - 1 || appointmentGoogleFoldersComboBox.SelectedIndex >= 1 && appointmentGoogleFoldersComboBox.SelectedIndex < appointmentGoogleFoldersComboBox.Items.Count)
                                                || !btSyncAppointments.Checked;

                //ToDo: Coloring doesn'T Work for these combos
                //setBgColor(contactFoldersComboBox, syncContactFolderIsValid);
                //setBgColor(noteFoldersComboBox, syncNoteFolderIsValid);
                //setBgColor(appointmentFoldersComboBox, syncAppointmentFolderIsValid);

                return syncContactFolderIsValid && /*syncNoteFolderIsValid &&*/ syncAppointmentFolderIsValid;
            }
        }

        private bool ValidSyncContactFolders
        {
            get
            {
                return (contactFoldersComboBox.SelectedIndex >= 1 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count)
                                                || !btSyncContacts.Checked;
            }
        }

        private bool ValidSyncAppointmentFolders
        {
            get
            {
                return (appointmentFoldersComboBox.SelectedIndex >= 1 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count)
                        && (appointmentGoogleFoldersComboBox.SelectedIndex == appointmentGoogleFoldersComboBox.Items.Count - 1 || appointmentGoogleFoldersComboBox.SelectedIndex >= 1 && appointmentGoogleFoldersComboBox.SelectedIndex < appointmentGoogleFoldersComboBox.Items.Count)
                                                || !btSyncAppointments.Checked;
            }
        }

        private bool ValidCredentials
        {
            get
            {
                bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
                //bool passwordIsValid = !string.IsNullOrEmpty(Password.Text.Trim());
                bool syncProfileIsValid = (cmbSyncProfile.SelectedIndex > 0 && cmbSyncProfile.SelectedIndex < cmbSyncProfile.Items.Count - 1);


                setBgColor(UserName, userNameIsValid);
                //setBgColor(Password, passwordIsValid);
                setBgColor(cmbSyncProfile, syncProfileIsValid);



                if (!userNameIsValid)
                    toolTip.SetToolTip(UserName, "User is of wrong format, should be full Google Mail address, e.g. user@googelmail.com");
                else
                    toolTip.SetToolTip(UserName, string.Empty);
                //if (!passwordIsValid)
                //    toolTip.SetToolTip(Password, "Password is empty, please provide your Google Mail password");
                //else
                //    toolTip.SetToolTip(Password, String.Empty);               


                return userNameIsValid &&
                       //passwordIsValid && 
                       syncProfileIsValid;
            }
        }

        private static void setBgColor(Control box, bool isValid)
        {
            if (!isValid)
                box.BackColor = Color.LightPink;
            else
                box.BackColor = Color.LightGreen;
        }

        private void syncButton_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private void Sync()
        {
            try
            {
                if (!ValidCredentials)
                    //return;
                    throw new Exception("E-Mail address and or Sync-Profile is incomplete or incorrect - Maybe a typo or no selection...");

                fillSyncFolderItems();

                if (!ValidSyncContactFolders)
                {
                    Logger.Log(@"contactFoldersComboBox.SelectedIndex: " + contactFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"contactFoldersComboBox.Items.Count: " + contactFoldersComboBox.Items.Count, EventType.Debug);
                    throw new Exception("At least one Outlook contact folder is not selected or invalid!");
                }

                if (!ValidSyncAppointmentFolders)
                {
                    Logger.Log(@"appointmentFoldersComboBox.SelectedIndex: " + appointmentFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"appointmentFoldersComboBox.Items.Count: " + appointmentFoldersComboBox.Items.Count, EventType.Debug);
                    Logger.Log(@"appointmentGoogleFoldersComboBox.SelectedIndex: " + appointmentGoogleFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"appointmentGoogleFoldersComboBox.Items.Count: " + appointmentGoogleFoldersComboBox.Items.Count, EventType.Debug);
                    throw new Exception("At least one Outlook appointment folder is not selected or invalid!");
                }

                //IconTimerSwitch(true);
                ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
                syncThread = new Thread(starter);
                syncThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                syncThread.CurrentUICulture = new CultureInfo("en-US");
                syncThread.Start();

                //if new version on sourceforge.net website than print an information to the log
                CheckVersion();

                // wait for thread to start
                for (int i = 0; !syncThread.IsAlive && i < 10; i++)
                    Thread.Sleep(1000);//DoNothing, until the thread was started, but only wait maximum 10 seconds


            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }

        [STAThread]
        private async void Sync_ThreadStarter()
        {
            //==>Instead of lock, use Interlocked to exit the code, if already another thread is calling the same
            bool won = false;

            try
            {

                won = Interlocked.CompareExchange(ref executing, 1, 0) == 0;
                if (won)
                {

                    TimerSwitch(false);

                    //if the contacts or notes folder has changed ==> Reset matches (to not delete contacts or notes on the one or other side)                
                    RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + SyncProfile);
                    string oldSyncContactsFolder = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
                    string oldSyncNotesFolder = regKeyAppRoot.GetValue(RegistrySyncNotesFolder) as string;
                    string oldSyncAppointmentsFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
                    string oldSyncAppointmentsGoogleFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;

                    //only reset notes if NotesFolder changed and reset contacts if ContactsFolder changed
                    //and only reset appointments, if either OutlookAppointmentsFolder changed (without changing Google at the same time) or GoogleAppointmentsFolder changed (without changing Outlook at the same time) (not chosen before means not changed)
                    bool syncContacts = !string.IsNullOrEmpty(oldSyncContactsFolder) && !oldSyncContactsFolder.Equals(this.syncContactsFolder) && btSyncContacts.Checked;
                    bool syncNotes = false; // !string.IsNullOrEmpty(oldSyncNotesFolder) && !oldSyncNotesFolder.Equals(this.syncNotesFolder) && btSyncNotes.Checked;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                    bool syncAppointments = !string.IsNullOrEmpty(oldSyncAppointmentsFolder) && !oldSyncAppointmentsFolder.Equals(this.syncAppointmentsFolder) && btSyncAppointments.Checked;
                    bool syncGoogleAppointments = !string.IsNullOrEmpty(this.syncAppointmentsGoogleFolder) && !this.syncAppointmentsGoogleFolder.Equals(oldSyncAppointmentsGoogleFolder) && btSyncAppointments.Checked;
                    if (syncContacts || /*syncNotes ||*/ syncAppointments && !syncGoogleAppointments || !syncAppointments && syncGoogleAppointments)
                    {
                        bool r = await ResetMatches(syncContacts, syncNotes, syncAppointments);
                        if (!r)
                            throw new Exception("Reset required but cancelled by user");
                    }

                    //Then save the Contacts and Notes Folders used at last sync
                    if (btSyncContacts.Checked)
                        regKeyAppRoot.SetValue(RegistrySyncContactsFolder, this.syncContactsFolder);
                    //ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                    //if (btSyncNotes.Checked)
                    //    regKeyAppRoot.SetValue(RegistrySyncNotesFolder, this.syncNotesFolder);
                    if (btSyncAppointments.Checked)
                    {
                        regKeyAppRoot.SetValue(RegistrySyncAppointmentsFolder, this.syncAppointmentsFolder);
                        if (string.IsNullOrEmpty(this.syncAppointmentsGoogleFolder) && !string.IsNullOrEmpty(oldSyncAppointmentsGoogleFolder))
                            this.syncAppointmentsGoogleFolder = oldSyncAppointmentsGoogleFolder;
                        if (!string.IsNullOrEmpty(this.syncAppointmentsGoogleFolder))
                            regKeyAppRoot.SetValue(RegistrySyncAppointmentsGoogleFolder, this.syncAppointmentsGoogleFolder);
                    }

                    SetLastSyncText("Syncing...");
                    notifyIcon.Text = Application.ProductName + "\nSyncing...";
                    //System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
                    //notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));                    
                    IconTimerSwitch(true);

                    SetFormEnabled(false);

                    if (sync == null)
                    {
                        sync = new Synchronizer();
                        sync.DuplicatesFound += new Synchronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                        sync.ErrorEncountered += new Synchronizer.ErrorNotificationHandler(OnErrorEncountered);
                        sync.TimeZoneChanges += new Synchronizer.TimeZoneNotificationHandler(OnTimeZoneChanges);
                    }

                    Logger.ClearLog();
                    SetSyncConsoleText("");
                    Logger.Log("Sync started (" + SyncProfile + ").", EventType.Information);
                    //SetSyncConsoleText(Logger.GetText());
                    sync.SyncProfile = SyncProfile;
                    Synchronizer.SyncContactsFolder = this.syncContactsFolder;
                    Synchronizer.SyncNotesFolder = this.syncNotesFolder;
                    Synchronizer.SyncAppointmentsFolder = this.syncAppointmentsFolder;
                    Synchronizer.SyncAppointmentsGoogleFolder = this.syncAppointmentsGoogleFolder;
                    Synchronizer.MonthsInPast = Convert.ToUInt16(this.pastMonthInterval.Value);
                    Synchronizer.MonthsInFuture = Convert.ToUInt16(this.futureMonthInterval.Value);
                    Synchronizer.Timezone = this.Timezone;

                    sync.SyncOption = syncOption;
                    sync.SyncDelete = btSyncDelete.Checked;
                    sync.PromptDelete = btPromptDelete.Checked && btSyncDelete.Checked;
                    sync.UseFileAs = chkUseFileAs.Checked;
                    sync.SyncNotes = false; // btSyncNotes.Checked;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
                    sync.SyncContacts = btSyncContacts.Checked;
                    sync.SyncAppointments = btSyncAppointments.Checked;
                    Synchronizer.SyncAppointmentsForceRTF = btSyncAppointmentsForceRTF.Checked;
                    Synchronizer.SyncContactsForceRTF = btSyncContactsForceRTF.Checked;

                    if (!sync.SyncContacts && !sync.SyncNotes && !sync.SyncAppointments)
                    {
                        SetLastSyncText("Sync failed.");
                        notifyIcon.Text = Application.ProductName + "\nSync failed";

                        string messageText = "Neither notes nor contacts nor appointments are switched on for syncing. Please choose at least one option. Sync aborted!";
                        //    Logger.Log(messageText, EventType.Error);
                        //    ShowForm();
                        //    ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        //    return;
                        //}

                        //if (sync.SyncAppointments && Syncronizer.Timezone == "")
                        //{
                        //    string messageText = "Please set your timezone before syncing your appointments! Sync aborted!";
                        Logger.Log(messageText, EventType.Error);
                        ShowForm();
                        ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        return;
                    }


                    sync.LoginToGoogle(UserName.Text);
                    sync.LoginToOutlook();

                    sync.Sync();

                    lastSync = DateTime.Now;
                    SetLastSyncText("Last synced at " + lastSync.ToString());

                    string message = string.Format("Sync complete.\r\n Synced:  {1} out of {0}.\r\n Deleted:  {2}.\r\n Skipped: {3}.\r\n Errors:    {4}.", sync.TotalCount, sync.SyncedCount, sync.DeletedCount, sync.SkippedCount, sync.ErrorCount);
                    Logger.Log(message, EventType.Information);

                    if (reportSyncResultCheckBox.Checked)
                    {
                        /*
                        notifyIcon.BalloonTipTitle = Application.ProductName;
                        notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);
                        */
                        ToolTipIcon icon;
                        if (sync.ErrorCount > 0)
                            icon = ToolTipIcon.Error;
                        else if (sync.SkippedCount > 0)
                            icon = ToolTipIcon.Warning;
                        else
                            icon = ToolTipIcon.Info;
                        /*notifyIcon.ShowBalloonTip(5000);
                        */
                        ShowBalloonToolTip(Application.ProductName,
                            string.Format("{0}. {1}", DateTime.Now, message),
                            icon,
                            5000, false);

                    }
                    string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                    if (sync.ErrorCount + sync.SkippedCount > 0)
                        toolTip += string.Format("\nWarnings: {0}.", sync.ErrorCount + sync.SkippedCount);
                    if (toolTip.Length >= 64)
                        toolTip = toolTip.Substring(0, 63);
                    notifyIcon.Text = toolTip;
                }
            }
            catch (Google.GData.Client.GDataRequestException ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                //string responseString = (null != ex.InnerException) ? ex.ResponseString : ex.Message;

                if (ex.InnerException is System.Net.WebException)
                {
                    string message = "Cannot connect to Google, please check for available internet connection and proxy settings if applicable: " + ex.InnerException.Message + "\r\n" + ex.ResponseString;
                    Logger.Log(message, EventType.Warning);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            catch (Google.GData.Client.InvalidCredentialsException)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                string message = "The credentials (Google Account username and/or password) are invalid, please correct them in the settings form before you sync again";
                Logger.Log(message, EventType.Error);
                ShowForm();
                ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);

            }
            catch (Exception ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";
                Logger.Log(ex, EventType.Debug);
                if (ex is COMException)
                {
                    string message = "Outlook exception, please assure that Outlook is running and not closed when syncing";
                    Logger.Log(message + ": " + ex.Message + "\r\n" + ex.StackTrace, EventType.Warning);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            finally
            {
                if (won)
                {
                    Interlocked.Exchange(ref executing, 0);
                    lastSync = DateTime.Now;
                    TimerSwitch(true);
                    SetFormEnabled(true);
                    if (sync != null)
                    {
                        sync.LogoffOutlook();
                        sync.LogoffGoogle();
                        sync = null;
                    }
                    IconTimerSwitch(false);
                }
            }
        }

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout, bool error)
        {
            //if user is active on workstation
            if (boolShowBalloonTip)
            {
                notifyIcon.BalloonTipTitle = title;
                notifyIcon.BalloonTipText = message;
                notifyIcon.BalloonTipIcon = icon;
                notifyIcon.ShowBalloonTip(timeout);
            }

            string iconText = title + ": " + message;
            if (!string.IsNullOrEmpty(iconText))
                notifyIcon.Text = (iconText).Substring(0, iconText.Length >= 63 ? 63 : iconText.Length);

            if (error)
                notifyIcon.Icon = this.IconError;
        }

        void Logger_LogUpdated(string Message)
        {
            AppendSyncConsoleText(Message);
        }

        void OnErrorEncountered(string title, Exception ex, EventType eventType)
        {
            // do not show ErrorHandler, as there may be multiple exceptions that would nag the user
            Logger.Log(ex.ToString(), EventType.Error);
            string message = string.Format("Error Saving Contact: {0}.\nPlease report complete ErrorMessage from Log to the Tracker\nat https://sourceforge.net/tracker/?group_id=369321", ex.Message);
            ShowBalloonToolTip(title, message, ToolTipIcon.Error, 5000, true);
            /*notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
            notifyIcon.ShowBalloonTip(5000);*/
        }

        void OnTimeZoneChanges(string timeZone)
        {
            if (appointmentTimezonesComboBox.InvokeRequired)
            {
                OnTimeZoneChangesCallback d = new OnTimeZoneChangesCallback(OnTimeZoneChanges);
                this.Invoke(d, new object[] { timeZone });
            }
            else
            {
                appointmentTimezonesComboBox.Text = timeZone;
            }
            Timezone = timeZone;
            Synchronizer.Timezone = timeZone;
        }

        void OnDuplicatesFound(string title, string message)
        {
            Logger.Log(message, EventType.Warning);
            ShowBalloonToolTip(title, message, ToolTipIcon.Warning, 5000, true);
            /*
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
			notifyIcon.ShowBalloonTip(5000);
             */
        }

        void OnNotificationReceived(string message)
        {
            SetLastSyncText(message);
        }

        public void SetFormEnabled(bool enabled)
        {
            if (this.InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(SetFormEnabled);
                this.Invoke(h, new object[] { enabled });
            }
            else
            {
                resetMatchesLinkLabel.Enabled = enabled;
                settingsGroupBox.Enabled = enabled;
                syncButton.Enabled = enabled;
                cancelButton.Enabled = !enabled;
            }
        }
        public void SetLastSyncText(string text)
        {
            if (this.InvokeRequired)
            {
                TextHandler h = new TextHandler(SetLastSyncText);
                this.Invoke(h, new object[] { text });
            }
            else
            {
                lastSyncLabel.Text = text;
            }
        }

        public void SetSyncConsoleText(string text)
        {
            if (this.InvokeRequired)
            {
                TextHandler h = new TextHandler(SetSyncConsoleText);
                this.Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }

        }
        public void AppendSyncConsoleText(string text)
        {
            if (this.InvokeRequired)
            {
                TextHandler h = new TextHandler(AppendSyncConsoleText);
                this.Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
        }
        public void TimerSwitch(bool value)
        {
            if (this.InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(TimerSwitch);
                this.Invoke(h, new object[] { value });
            }
            else
            {
                //If PC resumes or unlocks or is started, give him 5 minutes to recover everything before the sync starts
                if (lastSync <= DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0))
                    lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
                autoSyncInterval.Enabled = autoSyncCheckBox.Checked && value;
                syncTimer.Enabled = autoSyncCheckBox.Checked && value;
                nextSyncLabel.Visible = autoSyncCheckBox.Checked && value;
            }
        }



        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            //Logger.Log(m.Msg, EventType.Information);
            switch (m.Msg)
            {
                //System shutdown
                case NativeMethods.WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                /*case NativeMethods.WM_WTSSESSION_CHANGE:
                    {
                        int value = m.WParam.ToInt32();
                        //User Session locked
                        if (value == NativeMethods.WTS_SESSION_LOCK)
                        {
                            Console.WriteLine("Session Lock",EventType.Information);
                            //OnSessionLock();
                            boolShowBalloonTip = false; // Do something when locked
                        }
                        //User Session unlocked
                        else if (value == NativeMethods.WTS_SESSION_UNLOCK)
                        {
                            Console.WriteLine("Session Unlock", EventType.Information);
                            //OnSessionUnlock();
                            boolShowBalloonTip = true; // Do something when unlocked
                            TimerSwitch(true);
                        }
                     break;
                    }
                
                
                case NativeMethods.WM_POWERBROADCAST:
                    {
                        if (m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMEAUTOMATIC ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMECRITICAL ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESTANDBY ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESUSPEND ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED)
                        {                            
                            TimerSwitch(true);
                        }
                        else if (m.WParam.ToInt32() == NativeMethods.PBT_APMSUSPEND ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSUSPEND)
                        {
                            TimerSwitch(false);
                        }
                            

                        break;
                    }*/
                default:
                    break;
            }
            //Show Window from Tray
            if (m.Msg == NativeMethods.WM_GCSM_SHOWME)
                ShowForm();
            base.WndProc(ref m);
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!requestClose)
            {
                SaveSettings();
                e.Cancel = true;
            }
            HideForm();
        }
        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                cancellationTokenSource.Cancel();

                if (sync != null)
                    sync.LogoffOutlook();


                Logger.Log("Closed application.", EventType.Information);
                Logger.Close();

                SaveSettings();

                //unregister event handler
                SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                SystemEvents.PowerModeChanged -= SystemEvents_PowerModeSwitch;

                notifyIcon.Dispose();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }

        private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                int index = syncOptionBox.SelectedIndex;
                if (index == -1)
                    return;

                SetSyncOption(index);
            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }
        private void SetSyncOption(int index)
        {
            syncOption = (SyncOption)index;
            for (int i = 0; i < syncOptionBox.Items.Count; i++)
            {
                if (i == index)
                    syncOptionBox.SetItemCheckState(i, CheckState.Checked);
                else
                    syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        private void SettingsForm_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
                Hide();

        }

        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //if (WindowState == FormWindowState.Normal)
            //    HideForm();
            //else
            ShowForm();
        }

        private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            TimerSwitch(true);
        }

        private void syncTimer_Tick(object sender, EventArgs e)
        {

            TimeSpan syncTime = DateTime.Now - lastSync;
            TimeSpan limit = new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            if (syncTime < limit)
            {
                TimeSpan diff = limit - syncTime;
                string str = "Next sync in";
                if (diff.Hours != 0)
                    str += " " + diff.Hours + " h";
                if (diff.Minutes != 0 || diff.Hours != 0)
                    str += " " + diff.Minutes + " min";
                if (diff.Seconds != 0)
                    str += " " + diff.Seconds + " s";
                nextSyncLabel.Text = str;
            }
            else
            {
                Sync();
            }
        }

        private async void resetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            // force deactivation to show up
            Application.DoEvents();
            try
            {
                this.cancelButton.Enabled = false; //Cancel is only working for sync currently, not for reset
                await ResetMatches(btSyncContacts.Checked, false /*btSyncNotes.Checked*/, btSyncAppointments.Checked);//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
            }
            catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Logger.Log("Reset Matches failed", EventType.Error);
                ErrorHandler.Handle(ex);
            }
            finally
            {
                lastSync = DateTime.Now;
                TimerSwitch(true);
                SetFormEnabled(true);
                this.hideButton.Enabled = true;
                if (sync != null)
                {
                    sync.LogoffOutlook();
                    sync.LogoffGoogle();
                    sync = null;
                }
            }
        }

        private async Task<bool> ResetMatches(bool syncContacts, bool syncNotes, bool syncAppointments)
        {
            TimerSwitch(false);

            SetLastSyncText("Resetting matches...");
            notifyIcon.Text = Application.ProductName + "\nResetting matches...";

            fillSyncFolderItems();

            SetFormEnabled(false);

            if (sync == null)
            {
                sync = new Synchronizer();
            }

            Logger.ClearLog();
            SetSyncConsoleText("");
            Logger.Log("Reset Matches started  (" + SyncProfile + ").", EventType.Information);

            sync.SyncNotes = syncNotes;
            sync.SyncContacts = syncContacts;
            sync.SyncAppointments = syncAppointments;

            Synchronizer.SyncContactsFolder = syncContactsFolder;
            Synchronizer.SyncNotesFolder = syncNotesFolder;
            Synchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            Synchronizer.SyncAppointmentsGoogleFolder = syncAppointmentsGoogleFolder;
            sync.SyncProfile = SyncProfile;

            sync.LoginToGoogle(UserName.Text);
            sync.LoginToOutlook();


            if (sync.SyncAppointments)
            {
                bool deleteOutlookAppointments = false;
                bool deleteGoogleAppointments = false;

                switch (ShowDialog("Do you want to delete all Outlook Calendar entries?"))
                {
                    case DialogResult.Yes: deleteOutlookAppointments = true; break;
                    case DialogResult.No: deleteOutlookAppointments = false; break;
                    default: return false;
                }
                switch (ShowDialog("Do you want to delete all Google Calendar entries?"))
                {
                    case DialogResult.Yes: deleteGoogleAppointments = true; break;
                    case DialogResult.No: deleteGoogleAppointments = false; break;
                    default: return false;
                }

                Logger.Log("Resetting Google appointment matches...", EventType.Information);
                try
                {
                    await sync.ResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationTokenSource.Token);
                    sync.LoadAppointments();
                    sync.ResetOutlookAppointmentMatches(deleteOutlookAppointments);
                }
                catch (System.Threading.Tasks.TaskCanceledException)
                {
                    Logger.Log("Task cancelled by user.", EventType.Information);
                    sync.LoadAppointments();
                }
            }

            if (sync.SyncContacts)
            {
                sync.LoadContacts();
                sync.ResetContactMatches();
            }


            if (sync.SyncNotes)
            {
                sync.LoadNotes();
                sync.ResetNoteMatches();
            }
            lastSync = DateTime.Now;
            SetLastSyncText("Matches reset at " + lastSync.ToString());
            Logger.Log("Matches reset.", EventType.Information);

            return true;
        }


        public delegate DialogResult InvokeConflict(ConflictResolverForm conflictResolverForm);

        public DialogResult ShowConflictDialog(ConflictResolverForm conflictResolverForm)
        {
            if (this.InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeConflict(ShowConflictDialog), new object[] { conflictResolverForm });
            }
            else
            {
                DialogResult res = conflictResolverForm.ShowDialog(this);

                notifyIcon.Icon = this.Icon0;

                return res;

            }
        }
        private delegate void InvokeCallback();

        private void ShowForm()
        {
            if (this.InvokeRequired)
            {
                Invoke(new InvokeCallback(ShowForm));
            }
            else
            {
                FormWindowState oldState = WindowState;

                Show();
                Activate();
                WindowState = FormWindowState.Normal;
                fillSyncFolderItems();

                if (oldState != WindowState)
                    CheckVersion();
            }
        }

        private async void CheckVersion()
        {
            if (!NewVersionLinkLabel.Visible)
            {//Only check once, if new version is available

                try
                {
                    Cursor = Cursors.WaitCursor;
                    SuspendLayout();
                    //check for new version
                    if (NewVersionLinkLabel.LinkColor != Color.Red && await VersionInformation.isNewVersionAvailable(cancellationTokenSource.Token))
                    {
                        NewVersionLinkLabel.Visible = true;
                        NewVersionLinkLabel.LinkColor = Color.Red;
                        NewVersionLinkLabel.Text = "New Version of GCSM available on sf.net!";
                        notifyIcon.BalloonTipClicked += notifyIcon_BalloonTipClickedDownloadNewVersion;
                        ShowBalloonToolTip("New version available", "Click here to download", ToolTipIcon.Info, 20000, false);
                    }

                    NewVersionLinkLabel.Visible = true;

                }
                finally
                {
                    Cursor = Cursors.Default;
                    ResumeLayout();
                }
            }
        }

        private void notifyIcon_BalloonTipClickedDownloadNewVersion(object sender, System.EventArgs e)
        {
            Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            notifyIcon.BalloonTipClicked -= notifyIcon_BalloonTipClickedDownloadNewVersion;
        }

        private void HideForm()
        {
            WindowState = FormWindowState.Minimized;
            Hide();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm();
            this.Activate();
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            HideForm();
        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            requestClose = true;
            Close();
        }
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            using (AboutBox about = new AboutBox())
            {
                about.Show();
            }
        }
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(UserName.Text) ||
                string.IsNullOrEmpty(cmbSyncProfile.Text) /*||
                string.IsNullOrEmpty(contactFoldersComboBox.Text)*/ )
            {
                // this is the first load, show form
                ShowForm();
                UserName.Focus();
                ShowBalloonToolTip(Application.ProductName,
                        "Application started and visible in your PC's system tray, click on this balloon or the icon below to open the settings form and enter your Google credentials there.",
                        ToolTipIcon.Info,
                        5000, false);
            }
            else
                HideForm();
        }

        private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            string regKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
            try
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(regKey);
                if (runAtStartupCheckBox.Checked)
                {
                    // add to registry
                    regKeyAppRoot.SetValue("GoogleContactSync", "\"" + Application.ExecutablePath + "\"");
                }
                else
                {
                    // remove from registry
                    regKeyAppRoot.DeleteValue("GoogleContactSync");
                }
            }
            catch (Exception ex)
            {
                //if we can't write to that key, disable it... 
                runAtStartupCheckBox.Checked = false;
                runAtStartupCheckBox.Enabled = false;
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(new Exception(("Error saving 'Run program at startup' settings into Registry key '" + regKey + "' Error: " + ex.Message), ex));
            }
        }

        private void UserName_TextChanged(object sender, EventArgs e)
        {
            ValidateSyncButton();
        }

        private void ValidateSyncButton()
        {
            syncButton.Enabled = ValidCredentials && ValidSyncFolders;
        }

        private void deleteDuplicatesButton_Click(object sender, EventArgs e)
        {
            //DeleteDuplicatesForm f = new DeleteDuplicatesForm(_sync
        }

        private void Donate_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://sourceforge.net/project/project_donations.php?group_id=369321");
        }

        private void Donate_MouseEnter(object sender, EventArgs e)
        {
            Donate.BackColor = System.Drawing.Color.LightGray;
        }

        private void Donate_MouseLeave(object sender, EventArgs e)
        {
            Donate.BackColor = System.Drawing.Color.Transparent;
        }

        private void hideButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void proxySettingsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_proxy != null) _proxy.ShowDialog(this);
        }

        private void SettingsForm_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            ShowHelp();
        }

        private void SettingsForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            ShowHelp();
        }

        private static void ShowHelp()
        {
            // go to the page showing the help and howto instructions
            Process.Start("http://googlesyncmod.sourceforge.net/");
        }

        private void btSyncContacts_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on notes for syncing now).", "No sync switched on");
                btSyncAppointments.Checked = true;//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
            }
            contactFoldersComboBox.Visible = btSyncContacts.Checked;
            btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
        }

        //private void btSyncNotes_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (!btSyncContacts.Checked && !btSyncNotes.Checked && !btSyncAppointments.Checked)
        //    {
        //        MessageBox.Show("Neither notes nor contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on appointments for syncing now).", "No sync switched on");
        //        btSyncAppointments.Checked = true; //ToDo: Google Calendar Api v2 deprecated                
        //    }
        //    noteFoldersComboBox.Visible = btSyncNotes.Checked;
        //}

        private void btSyncAppointments_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on contacts for syncing now).", "No sync switched on");
                btSyncContacts.Checked = true;
            }
            appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = btSyncAppointments.Checked;
            this.labelTimezone.Visible = this.labelMonthsPast.Visible = this.labelMonthsFuture.Visible = this.btSyncAppointments.Checked;
            this.pastMonthInterval.Visible = this.futureMonthInterval.Visible = this.appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
            btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
        }

        private void cmbSyncProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if ((0 == comboBox.SelectedIndex) || (comboBox.SelectedIndex == (comboBox.Items.Count - 1)))
            {
                using (ConfigurationManagerForm _configs = new ConfigurationManagerForm())
                {

                    if (0 == comboBox.SelectedIndex && _configs != null)
                    {
                        SyncProfile = _configs.AddProfile();
                        ClearSettings();
                    }

                    if (comboBox.SelectedIndex == (comboBox.Items.Count - 1) && _configs != null)
                        _configs.ShowDialog(this);
                }
                fillSyncProfileItems();

                comboBox.Text = SyncProfile;
                SaveSettings();
            }
            if (comboBox.SelectedIndex < 0)
                MessageBox.Show("Please select Sync Profile.", "No sync switched on");
            else
            {
                //ClearSettings();
                LoadSettings(comboBox.Text);
                SyncProfile = comboBox.Text;
            }

            ValidateSyncButton();
        }

        private void contacFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Contacts folder you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                syncContactsFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((OutlookFolder)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncContactsFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }
            ValidateSyncButton();


        }

        private void noteFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Notes folder you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                syncNotesFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((OutlookFolder)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncNotesFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();


        }

        private void appointmentFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Appointments folder you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                syncAppointmentsFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((OutlookFolder)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncAppointmentsFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();


        }

        private void appointmentGoogleFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Google Calendar you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is GoogleCalendar)
            {
                syncAppointmentsGoogleFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((GoogleCalendar)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncAppointmentsGoogleFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();


        }

        private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            btPromptDelete.Visible = btSyncDelete.Checked;
            btPromptDelete.Checked = btSyncDelete.Checked;
        }

        private void pictureBoxExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Do you really want to exit " + Application.ProductName + "? This will also stop the service performing automatic synchronizaton in the background. If you only want to hide the settings form, use the 'Hide' Button instead.", "Exit " + Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                CancelButton_Click(sender, EventArgs.Empty); //Close running thread
                requestClose = true;
                Close();
            }
        }

        private void SystemEvents_PowerModeSwitch(Object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                TimerSwitch(false);
            }
            else if (e.Mode == PowerModes.Resume)
            {
                TimerSwitch(true);
            }
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                boolShowBalloonTip = false;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                boolShowBalloonTip = true;
                TimerSwitch(true);
            }
        }

        private void autoSyncInterval_ValueChanged(object sender, EventArgs e)
        {
            TimerSwitch(true);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
            KillSyncThread();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand, ControlThread = true)]
        private void KillSyncThread()
        {
            if (syncThread != null && syncThread.IsAlive)
                syncThread.Abort();
        }

        #region syncing icon
        public void IconTimerSwitch(bool value)
        {
            if (this.InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(IconTimerSwitch);
                this.Invoke(h, new object[] { value });
            }
            else
            {
                if (value) //Reset Icon to default icon as starting point for the syncing icon
                    notifyIcon.Icon = this.Icon0;
                iconTimer.Enabled = value;
            }
        }

        private void iconTimer_Tick(object sender, EventArgs e)
        {
            showNextIcon();
        }

        private void showNextIcon()
        {
            if (this.InvokeRequired)
            {
                IconHandler h = new IconHandler(showNextIcon);
                this.Invoke(h, new object[] { });
            }
            else
                this.notifyIcon.Icon = GetNextIcon(this.notifyIcon.Icon); ;
        }



        private Icon GetNextIcon(Icon currentIcon)
        {
            if (currentIcon == IconError) //Don't change the icon anymore, once an error occurred
                return IconError;
            if (currentIcon == Icon30)
                return Icon60;
            else if (currentIcon == Icon60)
                return Icon90;
            else if (currentIcon == Icon90)
                return Icon120;
            else if (currentIcon == Icon120)
                return Icon150;
            else if (currentIcon == Icon150)
                return Icon180;
            else if (currentIcon == Icon180)
                return Icon210;
            else if (currentIcon == Icon210)
                return Icon240;
            else if (currentIcon == Icon240)
                return Icon270;
            else if (currentIcon == Icon270)
                return Icon300;
            else if (currentIcon == Icon300)
                return Icon330;
            else if (currentIcon == Icon330)
                return Icon0;
            else
                return Icon30;
        }
        #endregion

        //private void futureMonthTextBox_Validating(object sender, CancelEventArgs e)
        //{
        //    ushort value;
        //    if (!ushort.TryParse(futureMonthTextBoxOld.Text, out value))
        //    {
        //        MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
        //        futureMonthTextBoxOld.Text = "0";
        //        e.Cancel = true;
        //    }

        //}

        //private void pastMonthTextBox_Validating(object sender, CancelEventArgs e)
        //{
        //    ushort value;
        //    if (!ushort.TryParse(pastMonthTextBoxOld.Text, out value))
        //    {
        //        MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
        //        pastMonthTextBoxOld.Text = "1";
        //        e.Cancel = true;
        //    }
        //}

        private void appointmentTimezonesComboBox_TextChanged(object sender, EventArgs e)
        {
            this.Timezone = appointmentTimezonesComboBox.Text;
        }

        private void linkLabelRevokeAuthentication_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Logger.Log("Trying to remove Authentication...", EventType.Information);
                FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);
                fDS.ClearAsync();
                Logger.Log("Removed Authentication...", EventType.Information);
            }
            catch (Exception ex)
            {
                Logger.Log(ex.ToString(), EventType.Error);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void appointmentGoogleFoldersComboBox_Enter(object sender, EventArgs e)
        {
            if (this.appointmentGoogleFoldersComboBox.DataSource == null ||
                this.appointmentGoogleFoldersComboBox.Items.Count <= 1)
            {
                Logger.Log("Loading Google Calendars...", EventType.Information);
                ArrayList googleAppointmentFolders = new ArrayList();

                this.appointmentGoogleFoldersComboBox.BeginUpdate();
                //this.appointmentGoogleFoldersComboBox.DataSource = null;

                Logger.Log("Loading Google Appointments folder...", EventType.Information);
                string defaultText = "    --- Select a Google Appointment folder ---";

                if (sync == null)
                    sync = new Synchronizer();

                sync.SyncAppointments = btSyncAppointments.Checked;
                sync.LoginToGoogle(UserName.Text);
                foreach (CalendarListEntry calendar in sync.calendarList)
                {
                    googleAppointmentFolders.Add(new GoogleCalendar(calendar.Summary, calendar.Id, calendar.Primary.HasValue ? calendar.Primary.Value : false));
                }

                if (googleAppointmentFolders != null) // && googleAppointmentFolders.Count > 0)
                {
                    googleAppointmentFolders.Sort();
                    googleAppointmentFolders.Insert(0, new GoogleCalendar(defaultText, defaultText, false));
                    this.appointmentGoogleFoldersComboBox.DataSource = googleAppointmentFolders;
                    this.appointmentGoogleFoldersComboBox.DisplayMember = "DisplayName";
                    this.appointmentGoogleFoldersComboBox.ValueMember = "FolderID";
                }
                this.appointmentGoogleFoldersComboBox.EndUpdate();
                this.appointmentGoogleFoldersComboBox.SelectedValue = defaultText;

                //Select Default Folder per Default
                foreach (GoogleCalendar folder in appointmentGoogleFoldersComboBox.Items)
                    if (folder.IsDefaultFolder)
                    {
                        this.appointmentGoogleFoldersComboBox.SelectedItem = folder;
                        break;
                    }
                Logger.Log("Loaded Google Calendars.", EventType.Information);
            }
        }

        private void autoSyncInterval_Enter(object sender, EventArgs e)
        {
            syncTimer.Enabled = false;
        }

        private void autoSyncInterval_Leave(object sender, EventArgs e)
        {
            //if (autoSyncInterval.Value == null)
            //{  //ToDo: Doesn'T work, if user deleted it, the Value is kept
            //    MessageBox.Show("No empty value allowed, set to minimum value: " + autoSyncInterval.Minimum);
            //    autoSyncInterval.Value = autoSyncInterval.Minimum;
            //}
            syncTimer.Enabled = true;
        }

        private void NewVersionLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (((LinkLabel)sender).LinkColor == Color.Red)
                Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            else
                Process.Start("https://sourceforge.net/projects/googlesyncmod");
        }


    }
}