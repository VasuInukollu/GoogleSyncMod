using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Util.Store;
using Google.Contacts;
using Google.GData.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal class Synchronizer : IDisposable
    {
        public const int OutlookUserPropertyMaxLength = 32;
        public const string OutlookUserPropertyTemplate = "g/con/{0}/";

        public static object _syncRoot = new object();
        internal static string UserName;

        public AppointmentsSynchronizer appointmentsSynchronizer = new AppointmentsSynchronizer();
        public ContactsSynchronizer contactsSynchronizer = new ContactsSynchronizer();

        public int TotalCount { get; set; }
        public int SyncedCount { get; private set; }
        public int DeletedCount { get; private set; }
        public int ErrorCount { get; private set; }
        public int SkippedCount { get; set; }
        public int SkippedCountNotMatches { get; set; }

        private ConflictResolution _ConflictResolution;
        public ConflictResolution ConflictResolution
        {
            get { return _ConflictResolution; }
            set { _ConflictResolution = value; appointmentsSynchronizer.ConflictResolution = value; contactsSynchronizer.ConflictResolution = value; }
        }

        private DeleteResolution _DeleteGoogleResolution;
        public DeleteResolution DeleteGoogleResolution
        {
            get { return _DeleteGoogleResolution; }
            set { _DeleteGoogleResolution = value; appointmentsSynchronizer.DeleteGoogleResolution = value; contactsSynchronizer.DeleteGoogleResolution = value; }
        }

        private DeleteResolution _DeleteOutlookResolution;
        public DeleteResolution DeleteOutlookResolution
        {
            get { return _DeleteOutlookResolution; }
            set { _DeleteOutlookResolution = value; appointmentsSynchronizer.DeleteOutlookResolution = value; contactsSynchronizer.DeleteOutlookResolution = value; }
        }

        public delegate void NotificationHandler(string message);

        public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
        public delegate void TimeZoneNotificationHandler(string timeZone);

        public event TimeZoneNotificationHandler TimeZoneChanges;

        private static Outlook.NameSpace _outlookNamespace;
        public static Outlook.NameSpace OutlookNameSpace
        {
            get
            {
                //Just create outlook instance again, in case the namespace is null
                CreateOutlookInstance();
                return _outlookNamespace;
            }
        }

        public static Outlook.Application OutlookApplication { get; private set; }

        private SyncOption _syncOption = SyncOption.MergeOutlookWins;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; appointmentsSynchronizer.SyncOption = value; contactsSynchronizer.SyncOption = value; }
        }

        private string _SyncProfile;
        public string SyncProfile
        {
            get { return _SyncProfile; }
            set { _SyncProfile = value; appointmentsSynchronizer.SyncProfile = value; contactsSynchronizer.SyncProfile = value; }
        }

        /// <summary>
        /// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
        /// </summary>
        /// 

        private bool _SyncDelete;
        public bool SyncDelete
        {
            get { return _SyncDelete; }
            set { _SyncDelete = value; appointmentsSynchronizer.SyncDelete = value; contactsSynchronizer.SyncDelete = value; }
        }

        public bool _PromptDelete;
        public bool PromptDelete
        {
            get { return _PromptDelete; }
            set { _PromptDelete = value; appointmentsSynchronizer.PromptDelete = value; contactsSynchronizer.PromptDelete = value; }
        }

        /// <summary>
        /// If true sync also contacts
        /// </summary>
        public bool SyncContacts { get; set; }

        /// <summary>
        /// If true sync also appointments (calendar)
        /// </summary>
        public bool SyncAppointments { get; set; }

        public void LoginToGoogle(string username)
        {
            Logger.Log("Connecting to Google...", EventType.Information);
            if (contactsSynchronizer.ContactsRequest == null && SyncContacts || appointmentsSynchronizer.EventRequest == null & SyncAppointments)
            {
                //OAuth2 for all services
                List<string> scopes = new List<string>();

                //Contacts-Scope
                scopes.Add("https://www.google.com/m8/feeds");

                //Calendar-Scope
                //Didn't work: scopes.Add("https://www.googleapis.com/auth/calendar");
                scopes.Add(CalendarService.Scope.Calendar);

                //take user credentials
                UserCredential credential;

                //load client secret from ressources
                byte[] jsonSecrets = Properties.Resources.client_secrets;

                //using (var stream = new FileStream(Application.StartupPath + "\\client_secrets.json", FileMode.Open, FileAccess.Read))
                using (var stream = new MemoryStream(jsonSecrets))
                {
                    FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);

                    GoogleClientSecrets clientSecrets = GoogleClientSecrets.Load(stream);

                    credential = GCSMOAuth2WebAuthorizationBroker.AuthorizeAsync(
                                    clientSecrets.Secrets,
                                    scopes.ToArray(),
                                    username,
                                    CancellationToken.None,
                                    fDS).
                                    Result;

                    var initializer = new Google.Apis.Services.BaseClientService.Initializer();
                    initializer.HttpClientInitializer = credential;

                    OAuth2Parameters parameters = new OAuth2Parameters
                    {
                        ClientId = clientSecrets.Secrets.ClientId,
                        ClientSecret = clientSecrets.Secrets.ClientSecret,

                        // Note: AccessToken is valid only for 60 minutes
                        AccessToken = credential.Token.AccessToken,
                        RefreshToken = credential.Token.RefreshToken
                    };
                    Logger.Log(Application.ProductName, EventType.Information);
                    RequestSettings settings = new RequestSettings(Application.ProductName, parameters);

                    if (SyncContacts)
                    {
                        //ContactsRequest = new ContactsRequest(rs);
                        contactsSynchronizer.ContactsRequest = new ContactsRequest(settings);
                    }

                    if (SyncAppointments)
                    {
                        appointmentsSynchronizer.ReadGoogleAppointmentConfig(initializer);
                    }
                }
            }

            UserName = username;

            int maxUserIdLength = OutlookUserPropertyMaxLength - (OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
            string userId = username;
            if (userId.Length > maxUserIdLength)
                userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.
            //Remove characters not allowed for Outlook user property names: []_#
            userId = userId.Replace("#", "").Replace("[", "").Replace("]", "").Replace("_", "");

            appointmentsSynchronizer.OutlookPropertyPrefix = string.Format(OutlookUserPropertyTemplate, userId);
            contactsSynchronizer.OutlookPropertyPrefix = string.Format(OutlookUserPropertyTemplate, userId);
        }

        public void LoginToOutlook()
        {
            Logger.Log("Connecting to Outlook...", EventType.Information);

            try
            {
                CreateOutlookInstance();
            }
            catch (Exception e)
            {
                if (!(e is COMException) && !(e is InvalidCastException))
                    throw;

                try
                {
                    // If outlook was closed/terminated inbetween, we will receive an Exception
                    // System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
                    // so recreate outlook instance
                    //And sometimes we we receive an Exception
                    // System.InvalidCastException 0x8001010E (RPC_E_WRONG_THREAD))
                    Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
                    /*OutlookApplication = new Outlook.Application();
                    _outlookNamespace = OutlookApplication.GetNamespace("mapi");
                    _outlookNamespace.Logon();*/
                    OutlookApplication = null;
                    _outlookNamespace = null;
                    CreateOutlookInstance();
                }
                catch (Exception ex)
                {
                    string message = "Cannot connect to Outlook.\r\nPlease restart " + Application.ProductName + " and try again. If error persists, please inform developers on OutlookForge.";
                    // Error again? We need full stacktrace, display it!
                    throw new Exception(message, ex);
                }
            }
        }

        private static void CreateOutlookApplication()
        {
            //Try to create new Outlook application 3 times, because mostly it fails the first time, if not yet running
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    // First try to get the running application in case Outlook is already started
                    try
                    {
                        OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                        break;  //Exit the for loop, if creating outlook application was successful
                    }
                    catch (COMException ex)
                    {
                        if (ex.ErrorCode == unchecked((int)0x80029c4a))
                        {
                            Logger.Log(ex, EventType.Debug);
                            throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                        }
                        // That failed - try to create a new application object, launching Outlook in the background
                        OutlookApplication = new Outlook.Application();
                        break;
                    }
                    catch (InvalidCastException ex)
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                    }
                    catch (Exception ex)
                    {
                        if (i == 2)
                        {
                            Logger.Log(ex, EventType.Debug);
                            throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                        }
                        else
                            Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == unchecked((int)0x80029c4a))
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                    }
                    if (i == 2)
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
                    }
                    else
                        Thread.Sleep(1000 * 10 * (i + 1));
                }
                catch (InvalidCastException ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (Exception ex)
                {
                    if (i == 2)
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
                    }
                    else
                        Thread.Sleep(1000 * 10 * (i + 1));
                }
            }
        }

        private static void CreateOutlookNamespace()
        {
            //Try to create new Outlook namespace 5 times, because mostly it fails the first time, if not yet running
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    _outlookNamespace = OutlookApplication.GetNamespace("MAPI");
                    break;  //Exit the for loop, if getting outlook namespace was successful
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == unchecked((int)0x80029c4a))
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                    }
                    if (i == 4)
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                    }
                    else
                    {
                        Logger.Log("Try: " + i, EventType.Debug);
                        Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
                catch (InvalidCastException ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (Exception ex)
                {
                    if (i == 4)
                    {
                        Logger.Log(ex, EventType.Debug);
                        throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                    }
                    else
                    {
                        Logger.Log("Try: " + i, EventType.Debug);
                        Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
            }
        }

        private static void CreateOutlookInstance()
        {
            if (OutlookApplication == null || _outlookNamespace == null)
            {
                CreateOutlookApplication();

                if (OutlookApplication == null)
                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");

                CreateOutlookNamespace();

                if (_outlookNamespace == null)
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
                else
                    Logger.Log("Connected to Outlook: " + VersionInformation.GetOutlookVersion(OutlookApplication), EventType.Debug);
            }

            //Just try to access the outlookNamespace to check, if it is still accessible, throws COMException, if not reachable 
            try
            {
                if (string.IsNullOrEmpty(ContactsSynchronizer.SyncContactsFolder))
                {
                    _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                }
                else
                {
                    _outlookNamespace.GetFolderFromID(ContactsSynchronizer.SyncContactsFolder);
                }
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == unchecked((int)0x80029c4a))
                {
                    Logger.Log(ex, EventType.Debug);
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                else if (ex.ErrorCode == unchecked((int)0x80040111)) //"The server is not available. Contact your administrator if this condition persists."
                {
                    try
                    {
                        Logger.Log("Trying to logon, 1st try", EventType.Debug);
                        _outlookNamespace.Logon("", "", false, false);
                        Logger.Log("1st try OK", EventType.Debug);
                    }
                    catch (Exception e1)
                    {
                        Logger.Log(e1, EventType.Debug);
                        try
                        {
                            Logger.Log("Trying to logon, 2nd try", EventType.Debug);
                            _outlookNamespace.Logon("", "", true, true);
                            Logger.Log("2nd try OK", EventType.Debug);
                        }
                        catch (Exception e2)
                        {
                            Logger.Log(e2, EventType.Debug);
                            throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", e2);
                        }
                    }
                }
                else
                {
                    Logger.Log(ex, EventType.Debug);
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                }
            }
        }

        public void LogoffOutlook()
        {
            try
            {
                Logger.Log("Disconnecting from Outlook...", EventType.Debug);
                if (_outlookNamespace != null)
                {
                    _outlookNamespace.Logoff();
                }
            }
            catch (Exception)
            {
                // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
                // so as outlook is closed anyways, we just ignore the exception here
            }
            finally
            {
                if (_outlookNamespace != null)
                {
                    Marshal.ReleaseComObject(_outlookNamespace);
                    _outlookNamespace = null;
                }
                if (OutlookApplication != null)
                {
                    Marshal.ReleaseComObject(OutlookApplication);
                    OutlookApplication = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                Logger.Log("Disconnected from Outlook", EventType.Debug);
            }
        }

        public void LogoffGoogle()
        {
            contactsSynchronizer.ContactsRequest = null;
        }

        public static Outlook.Items GetOutlookItems(Outlook.OlDefaultFolders outlookDefaultFolder, string syncFolder)
        {
            Outlook.MAPIFolder mapiFolder = null;
            if (string.IsNullOrEmpty(syncFolder))
            {
                mapiFolder = OutlookNameSpace.GetDefaultFolder(outlookDefaultFolder);
                if (mapiFolder == null)
                    throw new Exception("Error getting Default OutlookFolder: " + outlookDefaultFolder);
            }
            else
            {
                mapiFolder = OutlookNameSpace.GetFolderFromID(syncFolder);
                if (mapiFolder == null)
                    throw new Exception("Error getting OutlookFolder: " + syncFolder);
            }

            try
            {
                Outlook.Items items = mapiFolder.Items;
                if (items == null)
                    throw new Exception("Error getting Outlook items from OutlookFolder: " + mapiFolder.Name);
                else
                    return items;
            }
            finally
            {
                if (mapiFolder != null)
                {
                    Marshal.ReleaseComObject(mapiFolder);
                    mapiFolder = null;
                }
            }
        }

        private void LogSyncParams()
        {
            Logger.Log("Synchronization options:", EventType.Debug);
            Logger.Log("Profile: " + SyncProfile, EventType.Debug);
            Logger.Log("SyncOption: " + SyncOption, EventType.Debug);
            Logger.Log("SyncDelete: " + SyncDelete, EventType.Debug);
            Logger.Log("PromptDelete: " + PromptDelete, EventType.Debug);

            if (SyncContacts)
            {
                Logger.Log("Sync contacts", EventType.Debug);
                Logger.Log("SyncContactsFolder: " + ContactsSynchronizer.SyncContactsFolder, EventType.Debug);
                Logger.Log("SyncContactsForceRTF: " + ContactsSynchronizer.SyncContactsForceRTF, EventType.Debug);
                Logger.Log("UseFileAs: " + contactsSynchronizer.UseFileAs, EventType.Debug);
            }

            if (SyncAppointments)
            {
                Logger.Log("Sync appointments", EventType.Debug);
                Logger.Log("TimeMin: " + AppointmentsSynchronizer.TimeMin, EventType.Debug);
                Logger.Log("TimeMax: " + AppointmentsSynchronizer.TimeMax, EventType.Debug);
                Logger.Log("SyncAppointmentsFolder: " + AppointmentsSynchronizer.SyncAppointmentsFolder, EventType.Debug);
                Logger.Log("SyncAppointmentsGoogleFolder: " + AppointmentsSynchronizer.SyncAppointmentsGoogleFolder, EventType.Debug);
                Logger.Log("SyncAppointmentsForceRTF: " + AppointmentsSynchronizer.SyncAppointmentsForceRTF, EventType.Debug);
            }
        }

        public void Sync()
        {
            lock (_syncRoot)
            {
                try
                {
                    if (string.IsNullOrEmpty(SyncProfile))
                    {
                        Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                        return;
                    }

                    LogSyncParams();

                    SyncedCount = 0;
                    appointmentsSynchronizer.SyncedCount = 0;
                    contactsSynchronizer.SyncedCount = 0;
                    DeletedCount = 0;
                    appointmentsSynchronizer.DeletedCount = 0;
                    contactsSynchronizer.DeletedCount = 0;
                    ErrorCount = 0;
                    appointmentsSynchronizer.ErrorCount = 0;
                    contactsSynchronizer.ErrorCount = 0;
                    SkippedCount = 0;
                    appointmentsSynchronizer.SkippedCount = 0;
                    contactsSynchronizer.SkippedCount = 0;
                    SkippedCountNotMatches = 0;
                    appointmentsSynchronizer.SkippedCountNotMatches = 0;
                    contactsSynchronizer.SkippedCountNotMatches = 0;
                    ConflictResolution = ConflictResolution.Cancel;
                    DeleteGoogleResolution = DeleteResolution.Cancel;
                    DeleteOutlookResolution = DeleteResolution.Cancel;

                    if (SyncContacts)
                        contactsSynchronizer.MatchContacts();

                    if (SyncAppointments)
                    {
                        Logger.Log("Outlook default time zone: " + TimeZoneInfo.Local.Id, EventType.Information);
                        Logger.Log("Google default time zone: " + AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone, EventType.Information);
                        if (string.IsNullOrEmpty(AppointmentsSynchronizer.Timezone))
                        {
                            TimeZoneChanges?.Invoke(AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone);
                            Logger.Log("Timezone not configured, changing to default value from Google, it could be adjusted later in GUI.", EventType.Information);
                        }
                        else if (string.IsNullOrEmpty(AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone))
                        {
                            //Timezone was set, but some users do not have time zone set in Google
                            AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone = AppointmentsSynchronizer.Timezone;
                        }
                        AppointmentsSynchronizer.MappingBetweenTimeZonesRequired = false;
                        if (TimeZoneInfo.Local.Id != AppointmentSync.IanaToWindows(AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone))
                        {
                            AppointmentsSynchronizer.MappingBetweenTimeZonesRequired = true;
                            Logger.Log("Different time zones in Outlook (" + TimeZoneInfo.Local.Id + ") and Google (mapped to " + AppointmentSync.IanaToWindows(AppointmentsSynchronizer.SyncAppointmentsGoogleTimeZone) + ")", EventType.Warning);
                        }
                        appointmentsSynchronizer.MatchAppointments();
                    }

                    if (SyncContacts)
                    {
                        if (contactsSynchronizer.Contacts == null)
                            return;

                        TotalCount = contactsSynchronizer.Contacts.Count + SkippedCountNotMatches + appointmentsSynchronizer.SkippedCountNotMatches + contactsSynchronizer.SkippedCountNotMatches;

                        //Resolve Google duplicates from matches to be synced
                        contactsSynchronizer.ResolveDuplicateContacts(contactsSynchronizer.GoogleContactDuplicates);

                        //Remove Outlook duplicates from matches to be synced
                        if (contactsSynchronizer.OutlookContactDuplicates != null)
                        {
                            for (int i = contactsSynchronizer.OutlookContactDuplicates.Count - 1; i >= 0; i--)
                            {
                                ContactMatch match = contactsSynchronizer.OutlookContactDuplicates[i];
                                if (contactsSynchronizer.Contacts.Contains(match))
                                {
                                    //ToDo: If there has been a resolution for a duplicate above, there is still skipped increased, check how to distinguish
                                    SkippedCount++;
                                    contactsSynchronizer.Contacts.Remove(match);
                                }
                            }
                        }

                        Logger.Log("Syncing groups...", EventType.Information);
                        ContactsMatcher.SyncGroups(contactsSynchronizer);

                        Logger.Log("Syncing contacts...", EventType.Information);
                        ContactsMatcher.SyncContacts(contactsSynchronizer);

                        contactsSynchronizer.SaveContacts(contactsSynchronizer.Contacts);
                    }

                    if (SyncAppointments)
                    {
                        if (appointmentsSynchronizer.Appointments == null)
                            return;

                        TotalCount += appointmentsSynchronizer.Appointments.Count + SkippedCountNotMatches + appointmentsSynchronizer.SkippedCountNotMatches + contactsSynchronizer.SkippedCountNotMatches;

                        Logger.Log("Syncing appointments...", EventType.Information);
                        AppointmentsMatcher.SyncAppointments(appointmentsSynchronizer);

                        appointmentsSynchronizer.DeleteAppointments(appointmentsSynchronizer.Appointments);
                    }
                }
                finally
                {
                    if (contactsSynchronizer.OutlookContacts != null)
                    {
                        Marshal.ReleaseComObject(contactsSynchronizer.OutlookContacts);
                        contactsSynchronizer.OutlookContacts = null;
                    }
                    if (appointmentsSynchronizer.OutlookAppointments != null)
                    {
                        Marshal.ReleaseComObject(appointmentsSynchronizer.OutlookAppointments);
                        appointmentsSynchronizer.OutlookAppointments = null;
                    }
                    contactsSynchronizer.GoogleContacts = null;
                    appointmentsSynchronizer.GoogleAppointments = null;
                    contactsSynchronizer.OutlookContactDuplicates = null;
                    contactsSynchronizer.GoogleContactDuplicates = null;
                    contactsSynchronizer.GoogleGroups = null;
                    contactsSynchronizer.Contacts = null;
                    appointmentsSynchronizer.Appointments = null;
                }
            }
        }

        public void Dispose()
        {
            ((IDisposable)appointmentsSynchronizer).Dispose();
        }
    }

    internal enum SyncOption
    {
        MergePrompt,
        MergeOutlookWins,
        MergeGoogleWins,
        OutlookToGoogleOnly,
        GoogleToOutlookOnly,
    }
}
