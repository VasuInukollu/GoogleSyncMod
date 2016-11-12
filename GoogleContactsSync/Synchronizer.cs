using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Util.Store;
using Google.Contacts;
using Google.Documents;
using Google.GData.Client;
using Google.GData.Client.ResumableUpload;
using Google.GData.Contacts;
using Google.GData.Documents;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal class Synchronizer : IDisposable
    {
        public const int OutlookUserPropertyMaxLength = 32;
        public const string OutlookUserPropertyTemplate = "g/con/{0}/";
        internal const string myContactsGroup = "System Group: My Contacts";
        private static object _syncRoot = new object();
        internal static string UserName;

        public int TotalCount { get; private set; }
        public int SyncedCount { get; private set; }
        public int DeletedCount { get; private set; }
        public int ErrorCount { get; private set; }
        public int SkippedCount { get; set; }
        public int SkippedCountNotMatches { get; set; }
        public ConflictResolution ConflictResolution { get; set; }
        public DeleteResolution DeleteGoogleResolution { get; set; }
        public DeleteResolution DeleteOutlookResolution { get; set; }

        public delegate void NotificationHandler(string message);

        public delegate void DuplicatesFoundHandler(string title, string message);
        public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
        public delegate void TimeZoneNotificationHandler(string timeZone);

        public event DuplicatesFoundHandler DuplicatesFound;
        public event ErrorNotificationHandler ErrorEncountered;
        public event TimeZoneNotificationHandler TimeZoneChanges;

        public ContactsRequest ContactsRequest { get; private set; }

        private OAuth2Authenticator authenticator;

        public DocumentsRequest DocumentsRequest { get; private set; }
        public EventsResource EventRequest { get; private set; }

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
        public Outlook.Items OutlookContacts { get; private set; }
        public Outlook.Items OutlookNotes { get; private set; }
        public Outlook.Items OutlookAppointments { get; private set; }
        public Collection<ContactMatch> OutlookContactDuplicates { get; set; }
        public Collection<ContactMatch> GoogleContactDuplicates { get; set; }
        public Collection<Contact> GoogleContacts { get; private set; }
        public Collection<Document> GoogleNotes { get; private set; }
        private CalendarService CalendarRequest;
        public Collection<Google.Apis.Calendar.v3.Data.Event> GoogleAppointments { get; private set; }
        public Collection<Google.Apis.Calendar.v3.Data.Event> AllGoogleAppointments { get; private set; }
        public IList<CalendarListEntry> calendarList { get; private set; }
        public Collection<Group> GoogleGroups { get; set; }
        internal Document googleNotesFolder;
        public string OutlookPropertyPrefix { get; private set; }

        public string OutlookPropertyNameId
        {
            get { return OutlookPropertyPrefix + "id"; }
        }

        /*public string OutlookPropertyNameUpdated
        {
            get { return OutlookPropertyPrefix + "up"; }
        }*/

        public string OutlookPropertyNameSynced
        {
            get { return OutlookPropertyPrefix + "up"; }
        }

        private SyncOption _syncOption = SyncOption.MergeOutlookWins;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; }
        }

        public string SyncProfile { get; set; }
        public static string SyncContactsFolder { get; set; }
        public static string SyncNotesFolder { get; set; }
        public static string SyncAppointmentsFolder { get; set; }
        public static string SyncAppointmentsGoogleFolder { get; set; }
        public static string SyncAppointmentsGoogleTimeZone { get; set; }

        public static ushort MonthsInPast { get; set; }
        public static ushort MonthsInFuture { get; set; }
        public static string Timezone { get; set; }
        public static bool MappingBetweenTimeZonesRequired { get; set; }

        //private ConflictResolution? _conflictResolution;
        //public ConflictResolution? CResolution
        //{
        //    get { return _conflictResolution; }
        //    set { _conflictResolution = value; }
        //}

        public List<ContactMatch> Contacts { get; private set; }

        public List<NoteMatch> Notes { get; private set; }

        public List<AppointmentMatch> Appointments { get; private set; }

        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooMany = null;

        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooBig = null;

        private HashSet<string> ContactExtendedPropertiesToRemoveIfDuplicated = null;

        //private string _authToken;
        //public string AuthToken
        //{
        //    get
        //    {
        //        return _authToken;
        //    }
        //}

        /// <summary>
        /// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
        /// </summary>
        public bool SyncDelete { get; set; }
        public bool PromptDelete { get; set; }
        /// <summary>
        /// If true sync also notes
        /// </summary>
        public bool SyncNotes { get; set; }

        /// <summary>
        /// If true sync also contacts
        /// </summary>
        public bool SyncContacts { get; set; }
        public static bool SyncContactsForceRTF { get; set; }

        /// <summary>
        /// If true sync also appointments (calendar)
        /// </summary>
        public bool SyncAppointments { get; set; }
        public static bool SyncAppointmentsForceRTF { get; set; }
        /// <summary>
        /// if true, use Outlook's FileAs for Google Title/FullName. If false, use Outlook's Fullname
        /// </summary>
        public bool UseFileAs { get; set; }

        public void LoginToGoogle(string username)
        {
            Logger.Log("Connecting to Google...", EventType.Information);
            if (ContactsRequest == null && SyncContacts || DocumentsRequest == null && SyncNotes || EventRequest == null & SyncAppointments)
            {
                //OAuth2 for all services
                List<string> scopes = new List<string>();

                //Contacts-Scope
                scopes.Add("https://www.google.com/m8/feeds");
                //Notes-Scope
                //Obsolete, because no notes sync anymore: scopes.Add("https://docs.google.com/feeds/");
                //Didn'T work: scopes.Add("https://docs.googleusercontent.com/");
                //Didn'T work: scopes.Add("https://spreadsheets.google.com/feeds/");
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
                    RequestSettings settings = new RequestSettings(
                        Application.ProductName, parameters);

                    if (SyncContacts)
                    {
                        //ContactsRequest = new ContactsRequest(rs);
                        ContactsRequest = new ContactsRequest(settings);
                    }

                    //Obsolete, because no notes sync anymore:
                    if (SyncNotes)
                    {
                        //DocumentsRequest = new DocumentsRequest(rs);
                        DocumentsRequest = new DocumentsRequest(settings);

                        //Instantiate an Authenticator object according to your authentication, to use ResumableUploader
                        //GDataCredentials cred = new GDataCredentials(credential.Token.AccessToken);
                        //GOAuth2RequestFactory rf = new GOAuth2RequestFactory(null, Application.ProductName, parameters);
                        //DocumentsRequest.Service.RequestFactory = rf;

                        authenticator = new OAuth2Authenticator(Application.ProductName, parameters);
                    }
                    if (SyncAppointments)
                    {
                        //ContactsRequest = new Google.Contacts.ContactsRequest()

                        CalendarRequest = GoogleServices.CreateCalendarService(initializer);

                        //CalendarRequest.setUserCredentials(username, password);

                        calendarList = CalendarRequest.CalendarList.List().Execute().Items;

                        //Get Primary Calendar, if not set from outside
                        if (string.IsNullOrEmpty(SyncAppointmentsGoogleFolder))
                        {
                            foreach (var calendar in calendarList)
                            {
                                if (calendar.Primary != null && calendar.Primary.Value)
                                {
                                    SyncAppointmentsGoogleFolder = calendar.Id;
                                    SyncAppointmentsGoogleTimeZone = calendar.TimeZone;
                                    if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                                        Logger.Log("Empty Google time zone for calendar" + calendar.Id, EventType.Debug);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            bool found = false;
                            foreach (var calendar in calendarList)
                            {
                                if (calendar.Id == SyncAppointmentsGoogleFolder)
                                {
                                    SyncAppointmentsGoogleTimeZone = calendar.TimeZone;
                                    if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                                        Logger.Log("Empty Google time zone for calendar " + calendar.Id, EventType.Debug);
                                    else
                                        found = true;
                                    break;
                                }
                            }
                            if (!found)
                            {
                                Logger.Log("Cannot find calendar, id is " + SyncAppointmentsGoogleFolder, EventType.Warning);

                                Logger.Log("Listing calendars:", EventType.Debug);
                                foreach (var calendar in calendarList)
                                {
                                    if (calendar.Primary != null && calendar.Primary.Value)
                                    {
                                        Logger.Log("Id (primary): " + calendar.Id, EventType.Debug);
                                    }
                                    else
                                    {
                                        Logger.Log("Id: " + calendar.Id, EventType.Debug);
                                    }
                                }
                            }
                        }

                        if (SyncAppointmentsGoogleFolder == null)
                            throw new Exception("Google Calendar not defined (primary not found)");

                        //EventQuery query = new EventQuery("https://www.google.com/calendar/feeds/default/private/full");
                        //Old v2 approach: EventQuery query = new EventQuery("https://www.googleapis.com/calendar/v3/calendars/default/events");
                        EventRequest = CalendarRequest.Events;
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

            OutlookPropertyPrefix = string.Format(OutlookUserPropertyTemplate, userId);
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

            /*
            // Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
            // no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
            string profileName = "Outlook";
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
            {
                if (k != null)
                    profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
            }
            _outlookNamespace.Logon(profileName, null, true, false);*/

            //Just try to access the outlookNamespace to check, if it is still accessible, throws COMException, if not reachable 
            try
            {
                if (string.IsNullOrEmpty(SyncContactsFolder))
                {
                    _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                }
                else
                {
                    _outlookNamespace.GetFolderFromID(SyncContactsFolder);
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
                    Marshal.ReleaseComObject(_outlookNamespace);
                if (OutlookApplication != null)
                {
                    Marshal.ReleaseComObject(OutlookApplication);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                _outlookNamespace = null;
                OutlookApplication = null;
                Logger.Log("Disconnected from Outlook", EventType.Debug);
            }
        }

        public void LogoffGoogle()
        {
            ContactsRequest = null;
        }

        private void LoadOutlookContacts()
        {
            Logger.Log("Loading Outlook contacts...", EventType.Information);
            OutlookContacts = GetOutlookItems(Outlook.OlDefaultFolders.olFolderContacts, SyncContactsFolder);
            Logger.Log("Outlook Contacts Found: " + OutlookContacts.Count, EventType.Debug);
        }

        private void LoadOutlookNotes()
        {
            Logger.Log("Loading Outlook Notes...", EventType.Information);
            OutlookNotes = GetOutlookItems(Outlook.OlDefaultFolders.olFolderNotes, SyncNotesFolder);
            Logger.Log("Outlook Notes Found: " + OutlookNotes.Count, EventType.Debug);
        }

        private void LoadOutlookAppointments()
        {
            Logger.Log("Loading Outlook appointments...", EventType.Information);
            OutlookAppointments = GetOutlookItems(Outlook.OlDefaultFolders.olFolderCalendar, SyncAppointmentsFolder);
            Logger.Log("Outlook Appointments Found: " + OutlookAppointments.Count, EventType.Debug);
        }

        private Outlook.Items GetOutlookItems(Outlook.OlDefaultFolders outlookDefaultFolder, string syncFolder)
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

                //Outlook.MAPIFolder Folder = OutlookNameSpace.GetFolderFromID(_syncFolder);
                //if (Folder != null)
                //{
                //    for (int i = 1; i <= Folder.Folders.Count; i++)
                //    {
                //        Outlook.Folder subFolder = Folder.Folders[i] as Outlook.Folder;
                //        if ((Outlook.OlDefaultFolders.olFolderContacts == outlookDefaultFolder && Outlook.OlItemType.olContactItem == subFolder.DefaultItemType) ||
                //                 (Outlook.OlDefaultFolders.olFolderNotes == outlookDefaultFolder && Outlook.OlItemType.olNoteItem == subFolder.DefaultItemType) 
                //                )
                //        {
                //            mapiFolder = subFolder as Outlook.MAPIFolder;
                //        }
                //    }
                //}
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
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }


        ///// <summary>
        ///// Moves duplicates from OutlookContacts to OutlookContactDuplicates
        ///// </summary>
        //private void FilterOutlookContactDuplicates()
        //{
        //    OutlookContactDuplicates = new Collection<Outlook.ContactItem>();

        //    if (OutlookContacts.Count < 2)
        //        return;

        //    Outlook.ContactItem main, other;
        //    bool found = true;
        //    int index = 0;

        //    while (found)
        //    {
        //        found = false;

        //        for (int i = index; i <= OutlookContacts.Count - 1; i++)
        //        {
        //            main = OutlookContacts[i] as Outlook.ContactItem;

        //            // only look forward
        //            for (int j = i + 1; j <= OutlookContacts.Count; j++)
        //            {
        //                other = OutlookContacts[j] as Outlook.ContactItem;

        //                if (other.FileAs == main.FileAs &&
        //                    other.Email1Address == main.Email1Address)
        //                {
        //                    OutlookContactDuplicates.Add(other);
        //                    OutlookContacts.Remove(j);
        //                    found = true;
        //                    index = i;
        //                    break;
        //                }
        //            }
        //            if (found)
        //                break;
        //        }
        //    }
        //}

        private void LoadGoogleContacts()
        {
            LoadGoogleContacts(null);
            Logger.Log("Google Contacts Found: " + GoogleContacts.Count, EventType.Debug);
        }

        private Contact LoadGoogleContacts(AtomId id)
        {
            string message = "Error Loading Google Contacts. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Contact ret = null;
            try
            {
                if (id == null) // Only log, if not specific Google Contacts are searched                    
                    Logger.Log("Loading Google Contacts...", EventType.Information);

                GoogleContacts = new Collection<Contact>();

                ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;

                //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
                Group group = GetGoogleGroupByName(myContactsGroup);
                if (group != null)
                    query.Group = group.Id;

                //query.ShowDeleted = false;
                //query.OrderBy = "lastmodified";

                Feed<Contact> feed = ContactsRequest.Get<Contact>(query);

                while (feed != null)
                {
                    foreach (Contact a in feed.Entries)
                    {
                        GoogleContacts.Add(a);
                        if (id != null && id.Equals(a.ContactEntry.Id))
                            ret = a;
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = ContactsRequest.Get(feed, FeedRequestType.Next);
                }
            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            return ret;
        }
        private void LoadGoogleGroups()
        {
            string message = "Error Loading Google Groups. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            try
            {
                Logger.Log("Loading Google Groups...", EventType.Information);
                GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;
                //query.ShowDeleted = false;

                GoogleGroups = new Collection<Group>();

                Feed<Group> feed = ContactsRequest.Get<Group>(query);

                while (feed != null)
                {
                    foreach (Group a in feed.Entries)
                    {
                        GoogleGroups.Add(a);
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = ContactsRequest.Get(feed, FeedRequestType.Next);
                }

                ////Only for debugging or reset purpose: Delete all Gougle Groups:
                //for (int i = GoogleGroups.Count; i > 0;i-- )
                //    _googleService.Delete(GoogleGroups[i-1]);
            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }
        }

        private void LoadGoogleNotes()
        {
            LoadGoogleNotes(null, null);
            Logger.Log("Google Notes Found: " + GoogleNotes.Count, EventType.Debug);
        }

        internal Document LoadGoogleNotes(string folderUri, AtomId id)
        {
            string message = "Error Loading Google Notes. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Document ret = null;
            try
            {
                if (folderUri == null && id == null)
                {
                    // Only log, if not specific Google Notes are searched
                    Logger.Log("Loading Google Notes...", EventType.Information);
                    GoogleNotes = new Collection<Document>();
                }

                if (googleNotesFolder == null)
                    googleNotesFolder = GetOrCreateGoogleFolder(null, "Notes");//ToDo: Make the folder name Notes configurable in SettingsForm, for now hardcode to "Notes");

                if (folderUri == null)
                {
                    if (id == null)
                        folderUri = googleNotesFolder.DocumentEntry.Content.AbsoluteUri;
                    else //if newly created
                        folderUri = DocumentsRequest.BaseUri;
                }

                DocumentQuery query = new DocumentQuery(folderUri);
                query.Categories.Add(new QueryCategory(new AtomCategory("document")));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;

                //query.ShowDeleted = false;
                //query.OrderBy = "lastmodified";
                Feed<Document> feed = DocumentsRequest.Get<Document>(query);

                while (feed != null)
                {
                    foreach (Document a in feed.Entries)
                    {
                        if (id == null)
                            GoogleNotes.Add(a);
                        else if (id.Equals(a.DocumentEntry.Id))
                        {
                            ret = a;
                            return ret;
                        }
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = DocumentsRequest.Get(feed, FeedRequestType.Next);
                }

            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            return ret;
        }

        private void LoadGoogleAppointments()
        {
            Logger.Log("Loading Google appointments...", EventType.Information);
            LoadGoogleAppointments(null, MonthsInPast, MonthsInFuture, null, null);
            Logger.Log("Google Appointments Found: " + GoogleAppointments.Count, EventType.Debug);
        }

        /// <summary>
        /// Resets Google appointment matches.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        internal async Task ResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const int num_retries = 5;
            Logger.Log("Processing Google appointments.", EventType.Information);

            AllGoogleAppointments = null;
            GoogleAppointments = null;

            // First run batch updates, but since individual requests are not retried in case of any error rerun 
            // updates in single mode
            if (await BatchResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationToken))
            {
                // in case of error retry single updates five times
                for (var i = 1; i < num_retries; i++)
                {
                    if (!await SingleResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationToken))
                        break;
                }
            }

            Logger.Log("Finished all Google changes.", EventType.Information);
        }


        /// <summary>
        /// Resets Google appointment matches via single updates.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>If error occured.</returns>
        internal async Task<bool> SingleResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const string message = "Error resetting Google appointments.";
            try
            {
                var query = EventRequest.List(SyncAppointmentsGoogleFolder);
                string pageToken = null;

                if (MonthsInPast != 0)
                    query.TimeMin = DateTime.Now.AddMonths(-MonthsInPast);
                if (MonthsInFuture != 0)
                    query.TimeMax = DateTime.Now.AddMonths(MonthsInFuture);

                Logger.Log("Processing single updates.", EventType.Information);

                Events feed;
                bool gone_error = false;
                bool modified_error = false;

                do
                {
                    query.PageToken = pageToken;

                    //TODO (obelix30) - convert to Polly after retargeting to 4.5
                    try
                    {
                        feed = await query.ExecuteAsync(cancellationToken);
                    }
                    catch (Google.GoogleApiException ex)
                    {
                        if (GoogleServices.IsTransientError(ex.HttpStatusCode, ex.Error))
                        {
                            await Task.Delay(TimeSpan.FromMinutes(10), cancellationToken);
                            feed = await query.ExecuteAsync(cancellationToken);
                        }
                        else
                        {
                            throw new GDataRequestException(message, ex);
                        }
                    }

                    foreach (var a in feed.Items)
                    {
                        if (a.Id != null)
                        {
                            try
                            {
                                if (deleteGoogleAppointments)
                                {
                                    if (a.Status != "cancelled")
                                    {
                                        await EventRequest.Delete(SyncAppointmentsGoogleFolder, a.Id).ExecuteAsync(cancellationToken);
                                    }
                                }
                                else if (a.ExtendedProperties != null && a.ExtendedProperties.Shared != null && a.ExtendedProperties.Shared.ContainsKey("gos:oid:" + SyncProfile + ""))
                                {
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, a);
                                    if (a.Status != "cancelled")
                                    {
                                        await EventRequest.Update(a, SyncAppointmentsGoogleFolder, a.Id).ExecuteAsync(cancellationToken);
                                    }
                                }
                            }
                            catch (Google.GoogleApiException ex)
                            {
                                if (ex.HttpStatusCode == System.Net.HttpStatusCode.Gone)
                                {
                                    gone_error = true;
                                }
                                else if (ex.HttpStatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                {
                                    modified_error = true;
                                }
                                else
                                {
                                    throw new GDataRequestException("Exception", ex);
                                }
                            }
                        }
                    }
                    pageToken = feed.NextPageToken;
                }
                while (pageToken != null);

                if (modified_error)
                {
                    Logger.Log("Some Google appointments modified before update.", EventType.Debug);
                }
                if (gone_error)
                {
                    Logger.Log("Some Google appointments gone before deletion.", EventType.Debug);
                }
                return (gone_error || modified_error);
            }
            catch (System.Net.WebException ex)
            {
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }
        }


        /// <summary>
        /// Resets Google appointment matches via batch updates.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>If error occured.</returns>
        internal async Task<bool> BatchResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const string message = "Error updating Google appointments.";

            try
            {
                var query = EventRequest.List(SyncAppointmentsGoogleFolder);
                string pageToken = null;

                if (MonthsInPast != 0)
                    query.TimeMin = DateTime.Now.AddMonths(-MonthsInPast);
                if (MonthsInFuture != 0)
                    query.TimeMax = DateTime.Now.AddMonths(MonthsInFuture);

                Logger.Log("Processing batch updates.", EventType.Information);

                Events feed;
                var br = new BatchRequest(CalendarRequest);

                var events = new Dictionary<string, Google.Apis.Calendar.v3.Data.Event>();
                bool gone_error = false;
                bool modified_error = false;
                bool rate_error = false;
                bool current_batch_rate_error = false;

                int batches = 1;
                do
                {
                    query.PageToken = pageToken;

                    //TODO (obelix30) - check why sometimes exception happen like below,  we have custom backoff attached
                    //                    Google.GoogleApiException occurred
                    //User Rate Limit Exceeded[403]
                    //Errors[
                    //    Message[User Rate Limit Exceeded] Location[- ] Reason[userRateLimitExceeded] Domain[usageLimits]

                    //TODO (obelix30) - convert to Polly after retargeting to 4.5
                    try
                    {
                        feed = await query.ExecuteAsync(cancellationToken);
                    }
                    catch (Google.GoogleApiException ex)
                    {
                        if (GoogleServices.IsTransientError(ex.HttpStatusCode, ex.Error))
                        {
                            await Task.Delay(TimeSpan.FromMinutes(10), cancellationToken);
                            feed = await query.ExecuteAsync(cancellationToken);
                        }
                        else
                        {
                            throw new GDataRequestException(message, ex);
                        }
                    }

                    foreach (Google.Apis.Calendar.v3.Data.Event a in feed.Items)
                    {
                        if (a.Id != null && !events.ContainsKey(a.Id))
                        {
                            IClientServiceRequest r = null;
                            if (a.Status != "cancelled")
                            {
                                if (deleteGoogleAppointments)
                                {
                                    events.Add(a.Id, a);
                                    r = EventRequest.Delete(SyncAppointmentsGoogleFolder, a.Id);

                                }
                                else if (a.ExtendedProperties != null && a.ExtendedProperties.Shared != null && a.ExtendedProperties.Shared.ContainsKey("gos:oid:" + SyncProfile + ""))
                                {
                                    events.Add(a.Id, a);
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, a);
                                    r = EventRequest.Update(a, SyncAppointmentsGoogleFolder, a.Id);
                                }
                            }

                            if (r != null)
                            {
                                br.Queue<Google.Apis.Calendar.v3.Data.Event>(r, (content, error, ii, msg) =>
                                {
                                    if (error != null && msg != null)
                                    {
                                        if (msg.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                        {
                                            modified_error = true;
                                        }
                                        else if (msg.StatusCode == System.Net.HttpStatusCode.Gone)
                                        {
                                            gone_error = true;
                                        }
                                        else if (GoogleServices.IsTransientError(msg.StatusCode, error))
                                        {
                                            rate_error = true;
                                            current_batch_rate_error = true;
                                        }
                                        else
                                        {
                                            Logger.Log("Batch error: " + error.ToString(), EventType.Information);
                                        }
                                    }
                                });
                                if (br.Count >= GoogleServices.BatchRequestSize)
                                {
                                    if (current_batch_rate_error)
                                    {
                                        current_batch_rate_error = false;
                                        await Task.Delay(GoogleServices.BatchRequestBackoffDelay);
                                        Logger.Log("Back-Off waited " + GoogleServices.BatchRequestBackoffDelay + "ms before next retry...", EventType.Debug);

                                    }
                                    await br.ExecuteAsync(cancellationToken);
                                    // TODO(obelix30): https://github.com/google/google-api-dotnet-client/issues/725
                                    br = new BatchRequest(CalendarRequest);

                                    Logger.Log("Batch of Google changes finished (" + batches + ")", EventType.Information);
                                    batches++;
                                }
                            }
                        }
                    }
                    pageToken = feed.NextPageToken;
                }
                while (pageToken != null);

                if (br.Count > 0)
                {
                    await br.ExecuteAsync(cancellationToken);
                    Logger.Log("Batch of Google changes finished (" + batches + ")", EventType.Information);
                }
                if (modified_error)
                {
                    Logger.Log("Some Google appointment modified before update.", EventType.Debug);
                }
                if (gone_error)
                {
                    Logger.Log("Some Google appointment gone before deletion.", EventType.Debug);
                }
                if (rate_error)
                {
                    Logger.Log("Rate errors received.", EventType.Debug);
                }

                return (gone_error || modified_error || rate_error);
            }
            catch (System.Net.WebException ex)
            {
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }
        }

        internal Google.Apis.Calendar.v3.Data.Event LoadGoogleAppointments(string id, ushort restrictMonthsInPast, ushort restrictMonthsInFuture, DateTime? restrictStartTime, DateTime? restrictEndTime)
        {
            string message = "Error Loading Google appointments. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Google.Apis.Calendar.v3.Data.Event ret = null;
            try
            {

                GoogleAppointments = new Collection<Google.Apis.Calendar.v3.Data.Event>();

                var query = EventRequest.List(SyncAppointmentsGoogleFolder);

                string pageToken = null;
                //query.MaxResults = 256; //ToDo: Find a way to retrieve all appointments

                //Only Load events from month range, but onyl if not a distinct Google Appointment is searched for
                if (restrictMonthsInPast != 0)
                    query.TimeMin = DateTime.Now.AddMonths(-MonthsInPast);
                if (restrictStartTime != null && (query.TimeMin == default(DateTime) || restrictStartTime > query.TimeMin))
                    query.TimeMin = restrictStartTime.Value;
                if (restrictMonthsInFuture != 0)
                    query.TimeMax = DateTime.Now.AddMonths(MonthsInFuture);
                if (restrictEndTime != null && (query.TimeMax == default(DateTime) || restrictEndTime < query.TimeMax))
                    query.TimeMax = restrictEndTime.Value;

                //Doesn't work:
                //if (restrictStartDate != null)
                //    query.StartDate = restrictStartDate.Value;

                Events feed;

                do
                {
                    query.PageToken = pageToken;
                    feed = query.Execute();
                    foreach (Google.Apis.Calendar.v3.Data.Event a in feed.Items)
                    {
                        if ((a.RecurringEventId != null || !a.Status.Equals("cancelled")) &&
                            !GoogleAppointments.Contains(a) //ToDo: For an unknown reason, some appointments are duplicate in GoogleAppointments, therefore remove all duplicates before continuing  
                            )
                        {//only return not yet cancelled events (except for recurrence exceptions) and events not already in the list
                            GoogleAppointments.Add(a);
                            if (/*restrictStartDate == null && */id != null && id.Equals(a.Id))
                                ret = a;
                            //ToDo: Doesn't work for all recurrences
                            /*else if (restrictStartDate != null && id != null && a.RecurringEventId != null && a.Times.Count > 0 && restrictStartDate.Value.Date.Equals(a.Times[0].StartTime.Date))
                                if (id.Equals(new AtomId(id.AbsoluteUri.Substring(0, id.AbsoluteUri.LastIndexOf("/") + 1) + a.RecurringEventId.IdOriginal)))
                                    ret = a;*/
                        }
                        //else
                        //{
                        //    Logger.Log("Skipped Appointment because it was cancelled on Google side: " + a.Summary + " - " + GetTime(a), EventType.Information);
                        //SkippedCount++;
                        //}
                    }
                    pageToken = feed.NextPageToken;
                }
                while (pageToken != null);
            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            //Remember, if all Google Appointments have been loaded
            if (restrictMonthsInPast == 0 && restrictMonthsInFuture == 0 && restrictStartTime == null && restrictEndTime == null) //restrictStartDate == null)
                AllGoogleAppointments = GoogleAppointments;

            return ret;
        }

        /// <summary>
        /// Cleanup empty GoogleNotesFolders (for Outlook categories)
        /// </summary>
        internal void CleanUpGoogleCategories()
        {
            DocumentQuery query;
            Feed<Document> feed;
            List<Document> categoryFolders = GetGoogleGroups();

            if (categoryFolders != null)
            {
                foreach (Document categoryFolder in categoryFolders)
                {
                    query = new DocumentQuery(categoryFolder.DocumentEntry.Content.AbsoluteUri);
                    query.NumberToRetrieve = 256;
                    query.StartIndex = 0;

                    //query.ShowDeleted = false;
                    //query.OrderBy = "lastmodified";
                    feed = DocumentsRequest.Get<Document>(query);

                    bool isEmpty = true;
                    while (feed != null)
                    {
                        foreach (Document a in feed.Entries)
                        {
                            isEmpty = false;
                            break;
                        }
                        if (!isEmpty)
                            break;
                        query.StartIndex += query.NumberToRetrieve;
                        feed = DocumentsRequest.Get(feed, FeedRequestType.Next);
                    }

                    if (isEmpty)
                    {
                        DocumentsRequest.Delete(new Uri(DocumentsListQuery.documentsBaseUri + "/" + categoryFolder.ResourceId), categoryFolder.ETag);
                        Logger.Log("Deleted empty Google category folder: " + categoryFolder.Title, EventType.Information);
                    }
                }
            }
        }

        internal List<Document> GetGoogleGroups()
        {
            List<Document> categoryFolders;

            DocumentQuery query = new DocumentQuery(googleNotesFolder.DocumentEntry.Content.AbsoluteUri);
            query.Categories.Add(new QueryCategory(new AtomCategory("folder")));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;

            //query.ShowDeleted = false;
            //query.OrderBy = "lastmodified";
            Feed<Document> feed = DocumentsRequest.Get<Document>(query);
            categoryFolders = new List<Document>();

            while (feed != null)
            {
                foreach (Document a in feed.Entries)
                {
                    categoryFolders.Add(a);
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = DocumentsRequest.Get(feed, FeedRequestType.Next);
            }

            return categoryFolders;
        }

        private Document GetOrCreateGoogleFolder(Document parentFolder, string title)
        {
            Document ret = null;

            lock (this) //Synchronize the threads
            {
                ret = GetGoogleFolder(parentFolder, title, null);

                if (ret == null)
                {
                    ret = new Document();
                    ret.Type = Document.DocumentType.Folder;
                    //ret.Categories.Add(new AtomCategory("http://schemas.google.com/docs/2007#folder"));
                    ret.Title = title;
                    ret = SaveGoogleNote(parentFolder, ret, DocumentsRequest);
                }
            }

            return ret;
        }

        internal Document GetGoogleFolder(Document parentFolder, string title, string uri)
        {
            Document ret = null;

            //First get the Notes folder or create it, if not yet existing            
            DocumentQuery query = new DocumentQuery(DocumentsRequest.BaseUri);
            //Doesn't work, therefore used IsInFolder below: DocumentQuery query = new DocumentQuery((parentFolder == null) ? DocumentsRequest.BaseUri : parentFolder.DocumentEntryContent.AbsoluteUri);
            query.Categories.Add(new QueryCategory(new AtomCategory("folder")));
            if (!string.IsNullOrEmpty(title))
                query.Title = title;

            Feed<Document> feed = DocumentsRequest.Get<Document>(query);

            if (feed != null)
            {
                foreach (Document a in feed.Entries)
                {
                    if ((string.IsNullOrEmpty(uri) || a.Self == uri) &&
                        (parentFolder == null || IsInFolder(parentFolder, a)))
                    {
                        ret = a;
                        break;
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = DocumentsRequest.Get(feed, FeedRequestType.Next);
            }
            return ret;
        }
        /// <summary>
        /// Load the contacts from Google and Outlook
        /// </summary>
        public void LoadContacts()
        {
            LoadOutlookContacts();
            LoadGoogleGroups();
            LoadGoogleContacts();
            RemoveOutlookDuplicatedContacts();
            RemoveGoogleDuplicatedContacts();
        }

        public void LoadNotes()
        {
            LoadOutlookNotes();
            LoadGoogleNotes();
        }

        /// <summary>
        /// Remove duplicates from Google: two different Google appointments pointing to the same Outlook appointment.
        /// </summary>
        private void RemoveGoogleDuplicatedAppointments()
        {
            var appointments = new Dictionary<string, int>();

            //scan all Google appointments
            for (int i = 0; i < GoogleAppointments.Count; i++)
            {
                var e1 = GoogleAppointments[i];
                if (e1 == null)
                    continue;

                try
                {
                    string oid = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, e1);

                    //check if Google event is linked to Outlook appointment
                    if (string.IsNullOrEmpty(oid))
                        continue;

                    //check if there is already another Google event linked to the same Outlook appointment 
                    if (appointments.ContainsKey(oid))
                    {
                        var e2 = GoogleAppointments[appointments[oid]];
                        if (e2 == null)
                        {
                            appointments.Remove(oid);
                            continue;
                        }
                        var a = GetOutlookAppointmentById(oid);
                        if (a != null)
                        {
                            try
                            {
                                string gid = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, a);

                                //check to which Outlook appoinment Google event is linked
                                if (AppointmentPropertiesUtils.GetGoogleId(e1) == gid)
                                {
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e2);
                                    if (!string.IsNullOrEmpty(e2.Summary))
                                    {
                                        Logger.Log("Duplicated appointment: " + e2.Summary + ".", EventType.Debug);
                                    }
                                    appointments[oid] = i;
                                }
                                else if (AppointmentPropertiesUtils.GetGoogleId(e2) == gid)
                                {
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e1);
                                    if (!string.IsNullOrEmpty(e1.Summary))
                                    {
                                        Logger.Log("Duplicated appointment: " + e1.Summary + ".", EventType.Debug);
                                    }
                                }
                                else
                                {
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e1);
                                    AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e2);
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, a);
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(a);
                            }
                        }
                        else
                        {
                            //duplicated Google events found, but Outlook appointment does not exist
                            //so lets clean the link from Google events  
                            AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e1);
                            AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, e2);
                            appointments.Remove(oid);
                        }
                    }
                    else
                    {
                        appointments.Add(oid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some appointments throw exceptions
                    if (e1 != null && !string.IsNullOrEmpty(e1.Summary))
                        Logger.Log("Accessing Google appointment: " + e1.Summary + " threw and exception. Skipping: " + ex.Message, EventType.Debug);
                    else
                        Logger.Log("Accessing Google appointment threw and exception. Skipping: " + ex.Message, EventType.Debug);
                    continue;
                }
            }
        }

        /// <summary>
        /// Remove duplicates from Outlook: two different Outlook appointments pointing to the same Google appointment.
        /// Such situation typically happens when copy/paste'ing synchronized appointment in Outlook
        /// </summary>
        private void RemoveOutlookDuplicatedAppointments()
        {
            var appointments = new Dictionary<string, int>();

            //scan all appointments
            for (int i = 1; i <= OutlookAppointments.Count; i++)
            {
                Outlook.AppointmentItem ola1 = null;

                try
                {
                    ola1 = OutlookAppointments[i] as Outlook.AppointmentItem;
                    if (ola1 == null)
                        continue;

                    string gid = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, ola1);
                    //check if Outlook appointment is linked to Google event
                    if (string.IsNullOrEmpty(gid))
                        continue;

                    //check if there is already another Outlook appointment linked to the same Google event 
                    if (appointments.ContainsKey(gid))
                    {
                        var ola2 = OutlookAppointments[appointments[gid]] as Outlook.AppointmentItem;
                        if (ola2 == null)
                        {
                            appointments.Remove(gid);
                            continue;
                        }
                        try
                        {
                            var e = GetGoogleAppointmentById(gid);
                            if (e != null)
                            {
                                string oid = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, e);
                                //check to which Outlook appoinment Google event is linked
                                if (AppointmentPropertiesUtils.GetOutlookId(ola1) == oid)
                                {
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, ola2);
                                    if (!string.IsNullOrEmpty(ola2.Subject))
                                    {
                                        Logger.Log("Duplicated appointment: " + ola2.Subject + ".", EventType.Debug);
                                    }

                                    appointments[gid] = i;
                                }
                                else if (AppointmentPropertiesUtils.GetOutlookId(ola2) == oid)
                                {
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, ola1);
                                    if (!string.IsNullOrEmpty(ola1.Subject))
                                    {
                                        Logger.Log("Duplicated appointment: " + ola1.Subject + ".", EventType.Debug);
                                    }
                                }
                                else
                                {
                                    //duplicated Outlook appointments found, but Google event does not exist
                                    //so lets clean the link from Outlook appointments  
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, ola1);
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, ola2);
                                    appointments.Remove(gid);
                                }
                            }
                        }
                        finally
                        {
                            if (ola2 != null)
                                Marshal.ReleaseComObject(ola2);
                        }
                    }
                    else
                    {
                        appointments.Add(gid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some appointments throw exceptions
                    if (ola1 != null && !string.IsNullOrEmpty(ola1.Subject))
                        Logger.Log("Accessing Outlook appointment: " + ola1.Subject + " threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    else
                        Logger.Log("Accessing Outlook appointment threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    continue;
                }
                finally
                {
                    if (ola1 != null)
                        Marshal.ReleaseComObject(ola1);
                }
            }
        }

        /// <summary>
        /// Remove duplicates from Google: two different Google contacts pointing to the same Outlook contact.
        /// </summary>
        private void RemoveGoogleDuplicatedContacts()
        {
            var contacts = new Dictionary<string, int>();

            //scan all Google contacts
            for (int i = 0; i < GoogleContacts.Count; i++)
            {
                Contact c1 = GoogleContacts[i];
                if (c1 == null)
                    continue;

                try
                {
                    string oid = ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, c1);
                    //check if Google contact is linked to Outlook contact
                    if (string.IsNullOrEmpty(oid))
                        continue;

                    //check if there is already another Google contact linked to the same Outlook contact 
                    if (contacts.ContainsKey(oid))
                    {
                        var c2 = GoogleContacts[contacts[oid]];
                        if (c2 == null)
                        {
                            contacts.Remove(oid);
                            continue;
                        }

                        var a = GetOutlookContactById(oid);
                        if (a != null)
                        {
                            try
                            {
                                string gid = ContactPropertiesUtils.GetOutlookGoogleContactId(this, a);
                                //check to which Outlook contact Google contact is linked
                                if (ContactPropertiesUtils.GetGoogleId(c1) == gid)
                                {
                                    ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c2);
                                    if (!string.IsNullOrEmpty(c2.Title))
                                    {
                                        Logger.Log("Duplicated contact: " + c2.Title + ".", EventType.Debug);
                                    }
                                    contacts[oid] = i;
                                }
                                else if (ContactPropertiesUtils.GetGoogleId(c2) == gid)
                                {
                                    ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c1);
                                    if (!string.IsNullOrEmpty(c1.Title))
                                    {
                                        Logger.Log("Duplicated contact: " + c1.Title + ".", EventType.Debug);
                                    }
                                }
                                else
                                {
                                    ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c1);
                                    ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c2);
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, a);
                                }
                            }
                            finally
                            {
                                if (a != null)
                                    Marshal.ReleaseComObject(a);
                            }
                        }
                        else
                        {
                            //duplicated Google contacts found, but Outlook contact does not exist
                            //so lets clean the link from Google contacts
                            ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c1);
                            ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, c2);
                            contacts.Remove(oid);
                        }
                    }
                    else
                    {
                        contacts.Add(oid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    if (c1 != null && !string.IsNullOrEmpty(c1.Title))
                        Logger.Log("Accessing Google contact: " + c1.Title + " threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    else
                        Logger.Log("Accessing Google contact threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    continue;
                }
            }
        }

        /// <summary>
        /// Remove duplicates from Outlook: two different Outlook contacts pointing to the same Google contact.
        /// Such situation typically happens when copy/paste'ing synchronized contact in Outlook
        /// </summary>
        private void RemoveOutlookDuplicatedContacts()
        {
            var contacts = new Dictionary<string, int>();

            //scan all contacts
            for (int i = 1; i <= OutlookContacts.Count; i++)
            {
                Outlook.ContactItem olc1 = null;

                try
                {
                    olc1 = OutlookContacts[i] as Outlook.ContactItem;
                    if (olc1 == null)
                        continue;

                    string gid = ContactPropertiesUtils.GetOutlookGoogleContactId(this, olc1);
                    //check if Outlook contact  is linked to Google contact
                    if (string.IsNullOrEmpty(gid))
                        continue;

                    //check if there is already another Outlook contact linked to the same Google contact 
                    if (contacts.ContainsKey(gid))
                    {
                        var olc2 = OutlookContacts[contacts[gid]] as Outlook.ContactItem;
                        if (olc2 == null)
                        {
                            contacts.Remove(gid);
                            continue;
                        }
                        try
                        {
                            var c = GetGoogleContactById(gid);
                            if (c != null)
                            {
                                string oid = ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, c);
                                //check to which Outlook contact Google contact is linked
                                if (ContactPropertiesUtils.GetOutlookId(olc1) == oid)
                                {
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc2);
                                    if (!string.IsNullOrEmpty(olc2.FileAs))
                                    {
                                        Logger.Log("Duplicated contact: " + olc2.FileAs + ".", EventType.Debug);
                                    }
                                    contacts[oid] = i;
                                }
                                else if (ContactPropertiesUtils.GetOutlookId(olc2) == oid)
                                {
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc1);
                                    if (!string.IsNullOrEmpty(olc1.FileAs))
                                    {
                                        Logger.Log("Duplicated contact: " + olc1.FileAs + ".", EventType.Debug);
                                    }
                                }
                                else
                                {
                                    //duplicated Outlook contacts found, but Google contact does not exist
                                    //so lets clean the link from Outlook contacts  
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc1);
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc2);
                                    contacts.Remove(gid);
                                }
                            }
                            else
                            {
                                //duplicated Outlook contacts found, but Google contact does not exist
                                //so lets clean the link from Outlook contacts
                                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc1);
                                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, olc2);
                                contacts.Remove(gid);
                            }
                        }
                        finally
                        {
                            if (olc2 != null)
                                Marshal.ReleaseComObject(olc2);
                        }
                    }
                    else
                    {
                        contacts.Add(gid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    if (olc1 != null && !string.IsNullOrEmpty(olc1.FileAs))
                        Logger.Log("Accessing Outlook contact: " + olc1.FileAs + " threw and exception. Skipping: " + ex.Message, EventType.Debug);
                    else
                        Logger.Log("Accessing Outlook contact threw and exception. Skipping: " + ex.Message, EventType.Debug);
                    continue;
                }
                finally
                {
                    if (olc1 != null)
                        Marshal.ReleaseComObject(olc1);
                }
            }
        }
        public void LoadAppointments()
        {
            LoadOutlookAppointments();
            LoadGoogleAppointments();
            RemoveOutlookDuplicatedAppointments();
            RemoveGoogleDuplicatedAppointments();
        }

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void MatchContacts()
        {
            LoadContacts();

            DuplicateDataException duplicateDataException;
            Contacts = ContactsMatcher.MatchContacts(this, out duplicateDataException);
            if (duplicateDataException != null)
            {

                if (DuplicatesFound != null)
                    DuplicatesFound("Google duplicates found", duplicateDataException.Message);
                else
                    Logger.Log(duplicateDataException.Message, EventType.Warning);
            }

            Logger.Log("Contact Matches Found: " + Contacts.Count, EventType.Debug);
        }

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void MatchNotes()
        {
            LoadNotes();
            Notes = NotesMatcher.MatchNotes(this);
            /*DuplicateDataException duplicateDataException;
            _matches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
            if (duplicateDataException != null)
            {

                if (DuplicatesFound != null)
                    DuplicatesFound("Google duplicates found", duplicateDataException.Message);
                else
                    Logger.Log(duplicateDataException.Message, EventType.Warning);
            }*/
            Logger.Log("Note Matches Found: " + Notes.Count, EventType.Debug);
        }

        /// <summary>
        /// Load the appointments from Google and Outlook and match them
        /// </summary>
        public void MatchAppointments()
        {
            LoadAppointments();
            Appointments = AppointmentsMatcher.MatchAppointments(this);
            Logger.Log("Appointment Matches Found: " + Appointments.Count, EventType.Debug);
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

                    SyncedCount = 0;
                    DeletedCount = 0;
                    ErrorCount = 0;
                    SkippedCount = 0;
                    SkippedCountNotMatches = 0;
                    ConflictResolution = ConflictResolution.Cancel;
                    DeleteGoogleResolution = DeleteResolution.Cancel;
                    DeleteOutlookResolution = DeleteResolution.Cancel;

                    if (SyncContacts)
                        MatchContacts();

                    if (SyncNotes)
                        MatchNotes();

                    if (SyncAppointments)
                    {
                        Logger.Log("Outlook default time zone: " + TimeZoneInfo.Local.Id, EventType.Information);
                        Logger.Log("Google default time zone: " + SyncAppointmentsGoogleTimeZone, EventType.Information);
                        if (string.IsNullOrEmpty(Timezone))
                        {
                            TimeZoneChanges?.Invoke(SyncAppointmentsGoogleTimeZone);
                            Logger.Log("Timezone not configured, changing to default value from Google, it could be adjusted later in GUI.", EventType.Information);
                        }
                        else if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                        {
                            //Timezone was set, but some users do not have time zone set in Google
                            SyncAppointmentsGoogleTimeZone = Timezone;
                        }
                        MappingBetweenTimeZonesRequired = false;
                        if (TimeZoneInfo.Local.Id != AppointmentSync.IanaToWindows(SyncAppointmentsGoogleTimeZone))
                        {
                            MappingBetweenTimeZonesRequired = true;
                            Logger.Log("Different time zones in Outlook (" + TimeZoneInfo.Local.Id + ") and Google (mapped to " + AppointmentSync.IanaToWindows(SyncAppointmentsGoogleTimeZone) + ")", EventType.Warning);
                        }
                        MatchAppointments();
                    }

                    if (SyncContacts)
                    {
                        if (Contacts == null)
                            return;

                        TotalCount = Contacts.Count + SkippedCountNotMatches;

                        //Resolve Google duplicates from matches to be synced
                        ResolveDuplicateContacts(GoogleContactDuplicates);

                        //Remove Outlook duplicates from matches to be synced
                        if (OutlookContactDuplicates != null)
                        {
                            for (int i = OutlookContactDuplicates.Count - 1; i >= 0; i--)
                            {
                                ContactMatch match = OutlookContactDuplicates[i];
                                if (Contacts.Contains(match))
                                {
                                    //ToDo: If there has been a resolution for a duplicate above, there is still skipped increased, check how to distinguish
                                    SkippedCount++;
                                    Contacts.Remove(match);
                                }
                            }
                        }

                        Logger.Log("Syncing groups...", EventType.Information);
                        ContactsMatcher.SyncGroups(this);

                        Logger.Log("Syncing contacts...", EventType.Information);
                        ContactsMatcher.SyncContacts(this);

                        SaveContacts(Contacts);
                    }

                    if (SyncNotes)
                    {
                        if (Notes == null)
                            return;

                        TotalCount += Notes.Count + SkippedCountNotMatches;

                        Logger.Log("Syncing notes...", EventType.Information);
                        NotesMatcher.SyncNotes(this);

                        SaveNotes(Notes);

                        int timeout = 10;//seconds to wait for asynchronous upload
                        //Because notes are uploaded asynchonously, wait until all notes have been successfully uploaded
                        foreach (NoteMatch match in Notes)
                        {
                            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < timeout; i++)
                            {
                                Application.DoEvents();
                                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds
                                Application.DoEvents();
                            }

                            if (match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value)
                                Logger.Log("Asynchronous upload of note didn't finish within " + timeout + " seconds: " + match.GoogleNote.Title, EventType.Warning);
                        }

                        //Delete empty Google note folders
                        CleanUpGoogleCategories();
                    }

                    if (SyncAppointments)
                    {
                        if (Appointments == null)
                            return;

                        TotalCount += Appointments.Count + SkippedCountNotMatches; ;

                        Logger.Log("Syncing appointments...", EventType.Information);
                        AppointmentsMatcher.SyncAppointments(this);

                        DeleteAppointments(Appointments);
                    }

                }
                finally
                {
                    if (OutlookContacts != null)
                    {
                        Marshal.ReleaseComObject(OutlookContacts);
                        OutlookContacts = null;
                    }
                    if (OutlookNotes != null)
                    {
                        Marshal.ReleaseComObject(OutlookNotes);
                        OutlookNotes = null;
                    }
                    if (OutlookAppointments != null)
                    {
                        Marshal.ReleaseComObject(OutlookAppointments);
                        OutlookAppointments = null;
                    }
                    GoogleContacts = null;
                    GoogleNotes = null;
                    GoogleAppointments = null;
                    OutlookContactDuplicates = null;
                    GoogleContactDuplicates = null;
                    GoogleGroups = null;
                    Contacts = null;
                    Notes = null;
                    Appointments = null;
                }
            }
        }

        private void ResolveDuplicateContacts(Collection<ContactMatch> googleContactDuplicates)
        {
            if (googleContactDuplicates != null)
            {
                for (int i = googleContactDuplicates.Count - 1; i >= 0; i--)
                    ResolveDuplicateContact(googleContactDuplicates[i]);
            }
        }

        private void ResolveDuplicateContact(ContactMatch match)
        {
            if (Contacts.Contains(match))
            {
                if (_syncOption == SyncOption.MergePrompt)
                {
                    //For each OutlookDuplicate: Ask user for the GoogleContact to be synced with
                    for (int j = match.AllOutlookContactMatches.Count - 1; j >= 0 && match.AllGoogleContactMatches.Count > 0; j--)
                    {
                        OutlookContactInfo olci = match.AllOutlookContactMatches[j];
                        Outlook.ContactItem outlookContactItem = olci.GetOriginalItemFromOutlook();

                        try
                        {
                            Contact googleContact;
                            using (ConflictResolver r = new ConflictResolver())
                            {
                                switch (r.ResolveDuplicate(olci, match.AllGoogleContactMatches, out googleContact))
                                {
                                    case ConflictResolution.Skip:
                                    case ConflictResolution.SkipAlways: //Keep both entries and sync it to both sides
                                        match.AllGoogleContactMatches.Remove(googleContact);
                                        match.AllOutlookContactMatches.Remove(olci);
                                        Contacts.Add(new ContactMatch(null, googleContact));
                                        Contacts.Add(new ContactMatch(olci, null));
                                        break;
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google
                                        match.AllGoogleContactMatches.Remove(googleContact);
                                        match.AllOutlookContactMatches.Remove(olci);
                                        UpdateContact(outlookContactItem, googleContact);
                                        SaveContact(new ContactMatch(olci, googleContact));
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook
                                        match.AllGoogleContactMatches.Remove(googleContact);
                                        match.AllOutlookContactMatches.Remove(olci);
                                        UpdateContact(googleContact, outlookContactItem);
                                        SaveContact(new ContactMatch(olci, googleContact));
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }
                            }
                        }
                        finally
                        {
                            if (outlookContactItem != null)
                            {
                                Marshal.ReleaseComObject(outlookContactItem);
                                outlookContactItem = null;
                            }
                        }

                        //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                        if (match.AllOutlookContactMatches.Count == 0)
                            match.OutlookContact = null;
                        else
                            match.OutlookContact = match.AllOutlookContactMatches[0];
                    }
                }

                //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                if (match.AllGoogleContactMatches.Count == 0)
                    match.GoogleContact = null;
                else
                    match.GoogleContact = match.AllGoogleContactMatches[0];

                if (match.AllOutlookContactMatches.Count == 0)
                {
                    //If all OutlookContacts have been assigned by the users ==> Create one match for each remaining Google Contact to sync them to Outlook
                    Contacts.Remove(match);
                    foreach (Contact googleContact in match.AllGoogleContactMatches)
                        Contacts.Add(new ContactMatch(null, googleContact));
                }
                else if (match.AllGoogleContactMatches.Count == 0)
                {
                    //If all GoogleContacts have been assigned by the users ==> Create one match for each remaining Outlook Contact to sync them to Google
                    Contacts.Remove(match);
                    foreach (OutlookContactInfo outlookContact in match.AllOutlookContactMatches)
                        Contacts.Add(new ContactMatch(outlookContact, null));
                }
                else // if (match.AllGoogleContactMatches.Count > 1 ||
                //         match.AllOutlookContactMatches.Count > 1)
                {
                    SkippedCount++;
                    Contacts.Remove(match);
                }
                //else
                //{
                //    //If there remains a modified ContactMatch with only a single OutlookContact and GoogleContact
                //    //==>Remove all outlookContactDuplicates for this Outlook Contact to not remove it later from the Contacts to sync
                //    foreach (ContactMatch duplicate in OutlookContactDuplicates)
                //    {
                //        if (duplicate.OutlookContact.EntryID == match.OutlookContact.EntryID)
                //        {
                //            OutlookContactDuplicates.Remove(duplicate);
                //            break;
                //        }
                //    }
                //}
            }
        }

        public void DeleteAppointments(List<AppointmentMatch> appointments)
        {
            foreach (AppointmentMatch match in appointments)
            {
                try
                {
                    DeleteAppointment(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        string message = string.Format("Failed to synchronize appointment: {0}:\n{1}", match.OutlookAppointment != null ? match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start + ")" : match.GoogleAppointment.Summary + " - " + GetTime(match.GoogleAppointment), ex.Message);
                        Exception newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx, EventType.Error);
                    }
                    else
                        throw;
                }
            }
        }

        public static string GetTime(Google.Apis.Calendar.v3.Data.Event googleAppointment)
        {
            string ret = string.Empty;

            if (googleAppointment.Start != null && !string.IsNullOrEmpty(googleAppointment.Start.Date))
                ret += googleAppointment.Start.Date;
            else if (googleAppointment.Start != null && googleAppointment.Start.DateTime != null)
                ret += googleAppointment.Start.DateTime.Value.ToString();
            if (googleAppointment.Recurrence != null && googleAppointment.Recurrence.Count > 0)
                ret += " Recurrence"; //ToDo: Return Recurrence Start/End

            return ret;
        }

        public void DeleteAppointment(AppointmentMatch match)
        {
            if (match.GoogleAppointment != null && match.OutlookAppointment != null)
            {
                // Do nothing: Outlook appointments are not saved here anymore, they have already been saved and counted, just delete items

                ////bool googleChanged, outlookChanged;
                ////SaveAppointmentGroups(match, out googleChanged, out outlookChanged);
                //if (!match.GoogleAppointment.Saved)
                //{
                //    //Google appointment was modified. save.
                //    SyncedCount++;
                //    AppointmentPropertiesUtils.SetProperty(match.GoogleAppointment, Syncronizer.OutlookAppointmentsFolder, match.OutlookAppointment.EntryID);
                //    match.GoogleAppointment.Save();
                //    Logger.Log("Updated Google appointment from Outlook: \"" + match.GoogleAppointment.Summary + "\".", EventType.Information);
                //}

                //if (!match.OutlookAppointment.Saved)// || outlookChanged)
                //{
                //    //outlook appointment was modified. save.
                //    SyncedCount++;
                //    AppointmentPropertiesUtils.SetProperty(match.OutlookAppointment, Syncronizer.GoogleAppointmentsFolder, match.GoogleAppointment.EntryID);
                //    match.OutlookAppointment.Save();
                //    Logger.Log("Updated Outlook appointment from Google: \"" + match.OutlookAppointment.Subject + "\".", EventType.Information);
                //}                
            }
            else if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {
                if (match.OutlookAppointment.ItemProperties[OutlookPropertyNameId] != null)
                {
                    string name = match.OutlookAppointment.Subject;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook appointment because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook appointment because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // Google appointment was deleted, delete outlook appointment
                        Outlook.AppointmentItem item = match.OutlookAppointment;
                        //try
                        //{
                        string outlookAppointmentId = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, match.OutlookAppointment);
                        try
                        {
                            //First reset OutlookGoogleContactId to restore it later from trash
                            AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, item);
                            item.Save();
                        }
                        catch (Exception)
                        {
                            Logger.Log("Error resetting match for Outlook appointment: \"" + name + "\".", EventType.Warning);
                        }

                        item.Delete();

                        DeletedCount++;
                        Logger.Log("Deleted Outlook appointment: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(outlookContact);
                        //    outlookContact = null;
                        //}
                    }
                }
            }
            else if (match.GoogleAppointment != null && match.OutlookAppointment == null)
            {
                if (AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, match.GoogleAppointment) != null)
                {
                    string name = match.GoogleAppointment.Summary;
                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google appointment because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google appointment because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else if (match.GoogleAppointment.Status != "cancelled")
                    {
                        // outlook appointment was deleted, delete Google appointment
                        Google.Apis.Calendar.v3.Data.Event item = match.GoogleAppointment;
                        ////try
                        ////{
                        //string outlookAppointmentId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, match.GoogleAppointment);
                        //try
                        //{
                        //    //First reset OutlookGoogleContactId to restore it later from trash
                        //    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, item);
                        //    item.Save();
                        //}
                        //catch (Exception)
                        //{
                        //    Logger.Log("Error resetting match for Google appointment: \"" + name + "\".", EventType.Warning);
                        //}

                        EventRequest.Delete(SyncAppointmentsGoogleFolder, item.Id).Execute();

                        DeletedCount++;
                        Logger.Log("Deleted Google appointment: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(outlookContact);
                        //    outlookContact = null;
                        //}
                    }
                }
            }
            else
            {
                //TODO: ignore for now: 
                throw new ArgumentNullException("To save appointments, at least a GoogleAppointment or OutlookAppointment must be present.");
                //Logger.Log("Both Google and Outlook appointment: \"" + match.OutlookAppointment.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
            }
        }

        public void SaveContacts(List<ContactMatch> contacts)
        {
            foreach (ContactMatch match in contacts)
            {
                try
                {
                    SaveContact(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        string message = string.Format("Failed to synchronize contact: {0}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error on OutlookForge:\n{1}", match.OutlookContact != null ? match.OutlookContact.FileAs : match.GoogleContact.Title, ex.Message);
                        Exception newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx, EventType.Error);
                    }
                    else
                        throw;
                }
            }
        }

        public void SaveNotes(List<NoteMatch> notes)
        {
            foreach (NoteMatch match in notes)
            {
                try
                {
                    SaveNote(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        string message = string.Format("Failed to synchronize note: {0}.", match.OutlookNote.Subject);
                        Exception newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx, EventType.Error);
                    }
                    else
                        throw;
                }
            }
        }

        public void SaveContact(ContactMatch match)
        {
            if (match.GoogleContact != null && match.OutlookContact != null)
            {
                //bool googleChanged, outlookChanged;
                //SaveContactGroups(match, out googleChanged, out outlookChanged);
                if (match.GoogleContact.ContactEntry.Dirty || match.GoogleContact.ContactEntry.IsDirty())
                {
                    //google contact was modified. save.
                    SyncedCount++;
                    SaveGoogleContact(match);
                    Logger.Log("Updated Google contact from Outlook: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);
                }
            }
            else if (match.GoogleContact == null && match.OutlookContact != null)
            {
                if (match.OutlookContact.UserProperties.GoogleContactId != null)
                {
                    string name = match.OutlookContact.FileAs;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google contact was deleted, delete outlook contact
                        Outlook.ContactItem item = match.OutlookContact.GetOriginalItemFromOutlook();
                        try
                        {
                            try
                            {
                                //First reset OutlookGoogleContactId to restore it later from trash
                                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, item);
                                item.Save();
                            }
                            catch (Exception)
                            {
                                Logger.Log("Error resetting match for Outlook contact: \"" + name + "\".", EventType.Warning);
                            }

                            item.Delete();
                            DeletedCount++;
                            Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                            item = null;
                        }
                    }
                }
            }
            else if (match.GoogleContact != null && match.OutlookContact == null)
            {
                if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null)
                {
                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because of SyncOption " + _syncOption + ":" + ContactMatch.GetName(match.GoogleContact) + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because SyncDeletion is switched off :" + ContactMatch.GetName(match.GoogleContact) + ".", EventType.Information);
                    }
                    else
                    {
                        //commented oud, because it causes precondition failed error, if the ResetMatch is short before the Delete
                        //// peer outlook contact was deleted, delete google contact
                        //try
                        //{
                        //    //First reset GoogleOutlookContactId to restore it later from trash
                        //    match.GoogleContact = ResetMatch(match.GoogleContact);
                        //}
                        //catch (Exception)
                        //{
                        //    Logger.Log("Error resetting match for Google contact: \"" + ContactMatch.GetName(match.GoogleContact) + "\".", EventType.Warning);
                        //}

                        ContactsRequest.Delete(match.GoogleContact);
                        DeletedCount++;
                        Logger.Log("Deleted Google contact: \"" + ContactMatch.GetName(match.GoogleContact) + "\".", EventType.Information);
                    }
                }
            }
            else
            {
                //TODO: ignore for now: 
                throw new ArgumentNullException("To save contacts, at least a GoogleContacat or OutlookContact must be present.");
                //Logger.Log("Both Google and Outlook contact: \"" + match.OutlookContact.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
            }
        }

        public void SaveNote(NoteMatch match)
        {
            if (match.GoogleNote != null && match.OutlookNote != null)
            {
                //bool googleChanged, outlookChanged;
                //SaveNoteGroups(match, out googleChanged, out outlookChanged);
                if (match.GoogleNote.DocumentEntry.Dirty || match.GoogleNote.DocumentEntry.IsDirty())
                {
                    //google note was modified. save.
                    SyncedCount++;
                    SaveGoogleNote(match);
                    //Don't log here, because the DocumentsRequest uses async upload, log when async upload was successful
                    //Logger.Log("Updated Google note from Outlook: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
                }
                else if (!match.OutlookNote.Saved)// || outlookChanged) //If google note is saved above, Saving the OutlookNote not necessary anymore, because this will be done when updating NoteMatchId during saving the Google Note above
                {
                    //outlook note was modified. save.
                    SyncedCount++;
                    NotePropertiesUtils.SetOutlookGoogleNoteId(this, match.OutlookNote, match.GoogleNote);
                    match.OutlookNote.Save();
                    Logger.Log("Updated Outlook note from Google: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
                }

                // save photos
                //SaveNotePhotos(match);
            }
            else if (match.GoogleNote == null && match.OutlookNote != null)
            {
                if (match.OutlookNote.ItemProperties[OutlookPropertyNameId] != null)
                {
                    string name = match.OutlookNote.Subject;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook note because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google note was deleted, delete outlook note
                        Outlook.NoteItem item = match.OutlookNote;
                        //try
                        //{
                        string outlookNoteId = NotePropertiesUtils.GetOutlookGoogleNoteId(this, match.OutlookNote);
                        try
                        {
                            //First reset OutlookGoogleContactId to restore it later from trash
                            NotePropertiesUtils.ResetOutlookGoogleNoteId(this, item);
                            item.Save();
                        }
                        catch (Exception)
                        {
                            Logger.Log("Error resetting match for Outlook note: \"" + name + "\".", EventType.Warning);
                        }

                        item.Delete();
                        try
                        { //Delete also the according temporary NoteFile
                            File.Delete(NotePropertiesUtils.GetFileName(outlookNoteId, SyncProfile));
                        }
                        catch (Exception)
                        { }
                        DeletedCount++;
                        Logger.Log("Deleted Outlook note: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(item);
                        //    item = null;
                        //}
                    }
                }
            }
            else if (match.GoogleNote != null && match.OutlookNote == null)
            {
                if (NotePropertiesUtils.NoteFileExists(match.GoogleNote.Id, SyncProfile))
                {
                    string name = match.GoogleNote.Title;

                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google note because SyncDeletion is switched off :" + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer outlook note was deleted, delete google note
                        DocumentsRequest.Delete(new Uri(DocumentsListQuery.documentsBaseUri + "/" + match.GoogleNote.ResourceId), match.GoogleNote.ETag);
                        //DocumentsRequest.Service.Delete(match.GoogleNote.DocumentEntry); //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore I use the URI Delete above for now: "https://docs.google.com/feeds/default/private/full"
                        //DocumentsRequest.Delete(match.GoogleNote);

                        ////ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore the following workaround
                        //Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
                        //if (deletedNote != null)
                        //    DocumentsRequest.Delete(deletedNote);

                        try
                        {//Delete also the according temporary NoteFile
                            File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
                        }
                        catch (Exception)
                        { }

                        DeletedCount++;
                        Logger.Log("Deleted Google note: \"" + name + "\".", EventType.Information);
                    }
                }
            }
            else
            {
                //TODO: ignore for now: 
                throw new ArgumentNullException("To save notes, at least a GoogleContacat or OutlookNote must be present.");
                //Logger.Log("Both Google and Outlook note: \"" + match.OutlookNote.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
            }
        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public void UpdateAppointment(Outlook.AppointmentItem master, ref Google.Apis.Calendar.v3.Data.Event slave)
        {
            bool updated = false;
            if (slave.Creator != null && !AppointmentSync.IsOrganizer(slave.Creator.Email)) // && AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(this.SyncProfile, slave) != null)
            {
                //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment   
                switch (SyncOption)
                {
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite Outlook appointment
                        Logger.Log("Different Organizer found on Google, invitation maybe NOT sent by Outlook. Google appointment is overwriting Outlook because of SyncOption " + SyncOption + ": " + master.Subject + " - " + master.Start + ". ", EventType.Information);
                        UpdateAppointment(ref slave, master, null);
                        break;
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite Google appointment
                        Logger.Log("Different Organizer found on Google, invitation maybe NOT sent by Outlook, but Outlook appointment is overwriting Google because of SyncOption " + SyncOption + ": " + master.Subject + " - " + master.Start + ".", EventType.Information);
                        updated = true;
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (
                            //ConflictResolution != ConflictResolution.OutlookWinsAlways && //Shouldn't be used, because Google seems to be the master of the appointment
                            ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                ConflictResolution = r.Resolve("Cannot update appointment from Outlook to Google because different Organizer found on Google, invitation maybe NOT sent by Outlook: \"" + master.Subject + " - " + master.Start + "\". Do you want to update it back from Google to Outlook?", slave, master, this);
                            }
                        }
                        switch (ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Skip
                                SkippedCount++;
                                Logger.Log("Skipped Updating appointment from Outlook to Google because different Organizer found on Google, invitation maybe NOT sent by Outlook: \"" + master.Subject + " - " + master.Start + "\".", EventType.Information);
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook                           
                                UpdateAppointment(ref slave, master, null);
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google    
                                updated = true;
                                break;
                            default:
                                throw new ApplicationException("Cancelled");
                        }

                        break;
                }
            }
            else //Only update, if invitation was not sent on Google side or freshly created during this sync  
                updated = true;

            //if (master.Recipients.Count == 0 || 
            //    master.Organizer == null || 
            //    AppointmentSync.IsOrganizer(AppointmentSync.GetOrganizer(master), master)||
            //    slave.Id.Uri == null
            //    )
            //{//Only update, if this appointment was organized on Outlook side or freshly created during this sync

            if (updated)
            {
                AppointmentSync.UpdateAppointment(master, slave);

                if (slave.Creator == null || AppointmentSync.IsOrganizer(slave.Creator.Email))
                {
                    AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(SyncProfile, slave, master);
                    slave = SaveGoogleAppointment(slave);
                }

                //ToDo: Doesn'T work for newly created recurrence appointments before save, because Event.Reminder is throwing NullPointerException and Reminders cannot be initialized, therefore moved to after saving
                //if (slave.Recurrence != null && slave.Reminders != null)
                //{

                //    if (slave.Reminders.Overrides != null)
                //    {
                //        slave.Reminders.Overrides.Clear();
                //        if (master.ReminderSet)
                //        {
                //            var reminder = new Google.Apis.Calendar.v3.Data.EventReminder();
                //            reminder.Minutes = master.ReminderMinutesBeforeStart;
                //            if (reminder.Minutes > 40300)
                //            {
                //                //ToDo: Check real limit, currently 40300
                //                Logger.Log("Reminder Minutes to big (" + reminder.Minutes + "), set to maximum of 40300 minutes for appointment: " + master.Subject + " - " + master.Start, EventType.Warning);
                //                reminder.Minutes = 40300;
                //            }
                //            reminder.Method = "popup";
                //            slave.Reminders.Overrides.Add(reminder);
                //        }
                //    }
                //    slave = SaveGoogleAppointment(slave);
                //}

                AppointmentPropertiesUtils.SetOutlookGoogleAppointmentId(this, master, slave);
                master.Save();

                //After saving Google Appointment => also sync recurrence exceptions and save again
                if ((slave.Creator == null || AppointmentSync.IsOrganizer(slave.Creator.Email)) && master.IsRecurring && master.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster && AppointmentSync.UpdateRecurrenceExceptions(master, slave, this))
                {
                    slave = SaveGoogleAppointment(slave);
                }

                SyncedCount++;
                Logger.Log("Updated appointment from Outlook to Google: \"" + master.Subject + " - " + master.Start + "\".", EventType.Information);

                //}
                //else
                //{
                //    //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment
                //    SkippedCount++;
                //    //Logger.Log("Skipped Updating appointment from Outlook to Google because multiple recipients found and invitations NOT sent by Outlook: \"" + master.Subject + " - " + master.Start + "\".", EventType.Information);
                //    Logger.Log("Skipped Updating appointment from Outlook to Google because meeting was received by Outlook: \"" + master.Subject + " - " + master.Start + "\".", EventType.Information);
                //}
            }

        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public bool UpdateAppointment(ref Google.Apis.Calendar.v3.Data.Event master, Outlook.AppointmentItem slave, List<Google.Apis.Calendar.v3.Data.Event> googleAppointmentExceptions)
        {

            //if (master.Participants.Count > 1)
            //{
            //    bool organizerIsGoogle = AppointmentSync.IsOrganizer(AppointmentSync.GetOrganizer(master));

            //    if (organizerIsGoogle || AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, slave) == null)
            //    {//Only update, if this appointment was organized on Google side or freshly created during tis sync                    
            //        updated = true;
            //    }
            //    else
            //    {
            //        //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment
            //        SkippedCount++;
            //        Logger.Log("Skipped Updating appointment from Google to Outlook because multiple participants found and invitations NOT sent by Google: \"" + master.Summary + " - " + Syncronizer.GetTime(master) + "\".", EventType.Information);
            //    }
            //}
            //else                            
            //    updated = true;

            bool updated = false;
            if (slave.Recipients.Count > 1 && AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, slave) != null)
            {
                //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment   
                switch (SyncOption)
                {
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite Google appointment
                        Logger.Log("Multiple participants found, invitation maybe NOT sent by Google. Outlook appointment is overwriting Google because of SyncOption " + SyncOption + ": " + master.Summary + " - " + Synchronizer.GetTime(master) + ". ", EventType.Information);
                        UpdateAppointment(slave, ref master);
                        break;
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite outlook appointment
                        Logger.Log("Multiple participants found, invitation maybe NOT sent by Google, but Google appointment is overwriting Outlook because of SyncOption " + SyncOption + ": " + master.Summary + " - " + Synchronizer.GetTime(master) + ".", EventType.Information);
                        updated = true;
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (
                            //ConflictResolution != ConflictResolution.GoogleWinsAlways && //Shouldn't be used, because Outlook seems to be the master of the appointment
                            ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                            ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                ConflictResolution = r.Resolve("Cannot update appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + Synchronizer.GetTime(master) + "\". Do you want to update it back from Outlook to Google?", slave, master, this);
                            }
                        }
                        switch (ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Skip
                                SkippedCount++;
                                Logger.Log("Skipped Updating appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + Synchronizer.GetTime(master) + "\".", EventType.Information);
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google    
                                UpdateAppointment(slave, ref master);
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook
                                updated = true;
                                break;
                            default:
                                throw new ApplicationException("Cancelled");
                        }

                        break;
                }


                //if (MessageBox.Show("Cannot update appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + Syncronizer.GetTime(master) + "\". Do you want to update it back from Outlook to Google?", "Outlook appointment cannot be overwritten from Google", MessageBoxButtons.YesNo) == DialogResult.Yes)
                //    UpdateAppointment(slave, ref master);
                //else
                //    SkippedCount++;
                //    Logger.Log("Skipped Updating appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + Syncronizer.GetTime(master) + "\".", EventType.Information);
            }
            else //Only update, if invitation was not sent on Outlook side or freshly created during this sync  
                updated = true;

            if (updated)
            {
                AppointmentSync.UpdateAppointment(master, slave);
                AppointmentPropertiesUtils.SetOutlookGoogleAppointmentId(this, slave, master);
                try
                { //Try to save 2 times, because sometimes the first save fails with a COMException (Outlook aborted)
                    slave.Save();
                }
                catch (Exception)
                {
                    try
                    {
                        slave.Save();
                    }
                    catch (COMException ex)
                    {
                        Logger.Log("Error saving Outlook appointment: \"" + master.Summary + " - " + GetTime(master) + "\".\n" + ex.StackTrace, EventType.Warning);
                        return false;
                    }
                }

                if (master.Creator == null || AppointmentSync.IsOrganizer(master.Creator.Email))
                {
                    //only update Google, if I am the organizer, otherwise an error will be thrown
                    AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(SyncProfile, master, slave);
                    master = SaveGoogleAppointment(master);
                }

                SyncedCount++;
                Logger.Log("Updated appointment from Google to Outlook: \"" + master.Summary + " - " + GetTime(master) + "\".", EventType.Information);

                //After saving Outlook Appointment => also sync recurrence exceptions and increase SyncCount
                if (master.Recurrence != null && googleAppointmentExceptions != null && AppointmentSync.UpdateRecurrenceExceptions(googleAppointmentExceptions, slave, this))
                    SyncedCount++;
            }

            return true;
        }

        private void SaveOutlookContact(ref Contact googleContact, Outlook.ContactItem outlookContact)
        {
            ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
            outlookContact.Save();
            //Because Outlook automatically sets the EmailDisplayName to default value when the email is changed, update the emails again, to also sync the DisplayName
            ContactSync.SetEmails(googleContact, outlookContact);
            ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, googleContact, outlookContact);

            Contact updatedEntry = SaveGoogleContact(googleContact);
            //try
            //{
            //    updatedEntry = _googleService.Update(match.GoogleContact);
            //}
            //catch (GDataRequestException tmpEx)
            //{
            //    // check if it's the known HTCData problem, or if there is any invalid XML element or any unescaped XML sequence
            //    //if (tmpEx.ResponseString.Contains("HTCData") || tmpEx.ResponseString.Contains("&#39") || match.GoogleContact.Content.Contains("<"))
            //    //{
            //    //    bool wasDirty = match.GoogleContact.ContactEntry.Dirty;
            //    //    // XML escape the content
            //    //    match.GoogleContact.Content = EscapeXml(match.GoogleContact.Content);
            //    //    // set dirty to back, cause we don't want the changed content go back to Google without reason
            //    //    match.GoogleContact.ContactEntry.Content.Dirty = wasDirty;
            //    //    updatedEntry = _googleService.Update(match.GoogleContact);

            //    //}
            //    //else 
            //    if (!String.IsNullOrEmpty(tmpEx.ResponseString))
            //        throw new ApplicationException(tmpEx.ResponseString, tmpEx);
            //    else
            //        throw;
            //}            
            googleContact = updatedEntry;

            ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
            outlookContact.Save();
            SaveOutlookPhoto(googleContact, outlookContact);
        }

        private static string EscapeXml(string xml)
        {
            return System.Security.SecurityElement.Escape(xml);
        }

        public void SaveGoogleContact(ContactMatch match)
        {
            Outlook.ContactItem outlookContactItem = match.OutlookContact.GetOriginalItemFromOutlook();
            try
            {
                ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, outlookContactItem);
                match.GoogleContact = SaveGoogleContact(match.GoogleContact);
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactItem, match.GoogleContact);
                outlookContactItem.Save();

                //Now save the Photo
                SaveGooglePhoto(match, outlookContactItem);

            }
            finally
            {
                Marshal.ReleaseComObject(outlookContactItem);
                outlookContactItem = null;
            }
        }

        public void SaveGoogleNote(NoteMatch match)
        {
            Outlook.NoteItem outlookNoteItem = match.OutlookNote;
            //try
            //{  

            //ToDo: Somewhow, the content is not uploaded to Google, only an empty document                
            //match.GoogleNote = SaveGoogleNote(match.GoogleNote);

            //New approach how to update an existing document: https://developers.google.com/google-apps/documents-list/#updatingchanging_documents_and_files
            // Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();
            // Set the handlers for the completion and progress events                  
            //uploader.AsyncOperationProgress += new AsyncOperationProgressEventHandler(OnProgress);

            //ToDo: Therefoe I use DocumentService.UploadDocument instead and move it to the NotesFolder
            string oldOutlookGoogleNoteId = NotePropertiesUtils.GetOutlookGoogleNoteId(this, outlookNoteItem);
            if (match.GoogleNote.DocumentEntry.Id.Uri != null)
            {
                //DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + match.GoogleNote.ResourceId), match.GoogleNote.ETag);
                ////DocumentsRequest.Delete(match.GoogleNote); //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder
                //NotePropertiesUtils.ResetOutlookGoogleNoteId(this, outlookNoteItem);                                        

                ////ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder
                //Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
                //if (deletedNote != null)
                //    DocumentsRequest.Delete(deletedNote);

                // Start the update process.  
                uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteUpdated);
                uploader.UpdateAsync(authenticator, match.GoogleNote.DocumentEntry, match);

                //uploader.Update(_authenticator, match.GoogleNote.DocumentEntry);
            }
            else
            {
                uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteCreated);
                CreateGoogleNote(match.GoogleNote, match, DocumentsRequest, uploader, authenticator);
            }

            match.AsyncUpdateCompleted = false;

            //Google.GData.Documents.DocumentEntry entry = DocumentsRequest.Service.UploadDocument(NotePropertiesUtils.GetFileName(outlookNoteItem.EntryID, SyncProfile), match.GoogleNote.Title.Replace(":", String.Empty));                               
            //Document newNote = LoadGoogleNotes(entry.Id);
            //match.GoogleNote = DocumentsRequest.MoveDocumentTo(GoogleNotesFolder, newNote);

            //First delete old temporary file, because it was saved with old GoogleNoteID, because every sync to Google becomes a new ID, because updateMedia doesn't work
            //File.Delete(NotePropertiesUtils.GetFileName(oldOutlookGoogleNoteId, SyncProfile));
            //UpdateNoteMatchId(match);
            //}
            //finally
            //{
            //    Marshal.ReleaseComObject(outlookNoteItem);
            //    outlookNoteItem = null;
            //}
        }

        public static void CreateGoogleNote(/*Document parentFolder, */Document googleNote, object UserData, DocumentsRequest documentsRequest, ResumableUploader uploader, OAuth2Authenticator authenticator)
        {
            // Define the resumable upload link      
            Uri createUploadUrl = new Uri("https://docs.google.com/feeds/upload/create-session/default/private/full");
            //Uri createUploadUrl = new Uri(GoogleNotesFolder.AtomEntry.EditUri.ToString()); 
            AtomLink link = new AtomLink(createUploadUrl.AbsoluteUri);
            link.Rel = ResumableUploader.CreateMediaRelation;
            googleNote.DocumentEntry.Links.Add(link);
            //if (parentFolder != null)
            //    googleNote.DocumentEntry.ParentFolders.Add(new AtomLink(parentFolder.DocumentEntry.SelfUri.ToString()));
            // Set the service to be used to parse the returned entry 
            googleNote.DocumentEntry.Service = documentsRequest.Service;
            // Start the upload process   
            //uploader.InsertAsync(_authenticator, match.GoogleNote.DocumentEntry, new object());
            uploader.InsertAsync(authenticator, googleNote.DocumentEntry, UserData);
        }

        private void UpdateNoteMatchId(NoteMatch match)
        {
            NotePropertiesUtils.SetOutlookGoogleNoteId(this, match.OutlookNote, match.GoogleNote);
            match.OutlookNote.Save();

            //As GoogleDocuments don't have UserProperties, we have to use the file to check, if Note was already synced or not
            File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
            File.Move(NotePropertiesUtils.GetFileName(match.OutlookNote.EntryID, SyncProfile), NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
        }

        private void OnGoogleNoteCreated(object sender, AsyncOperationCompletedEventArgs e)
        {
            MoveGoogleNote(e.Entry as DocumentEntry, e.UserState as NoteMatch, true, e.Error, e.Cancelled);
        }

        private void OnGoogleNoteUpdated(object sender, AsyncOperationCompletedEventArgs e)
        {
            MoveGoogleNote(e.Entry as DocumentEntry, e.UserState as NoteMatch, false, e.Error, e.Cancelled);
        }
        private void MoveGoogleNote(DocumentEntry entry, NoteMatch match, bool create, Exception ex, bool cancelled)
        {
            if (ex != null)
            {
                ErrorHandler.Handle(new Exception("Google Note couldn't be " + (create ? "created" : "updated") + " :" + entry == null ? null : entry.Summary.ToString(), ex));
                return;
            }

            if (cancelled || entry == null)
            {
                ErrorHandler.Handle(new Exception("Google Note " + (create ? "creation" : "update") + " was cancelled: " + entry == null ? null : entry.Summary.ToString()));
                return;
            }

            //Get updated Google Note
            Document newNote = LoadGoogleNotes(null, entry.Id);
            match.GoogleNote = newNote;

            //Doesn't work because My Drive is not listed as parent folder: Remove all parent folders except for the Notes subfolder
            //if (create)
            //{
            //    foreach (string parentFolder in newNote.ParentFolders)
            //        if (parentFolder != googleNotesFolder.Self)
            //            DocumentsRequest.Delete(new Uri(googleNotesFolder.DocumentEntry.Content.AbsoluteUri + "/" + newNote.ResourceId),newNote.ETag);
            //}

            //first delete the note from all categories, the still valid categories are assigned again later           
            foreach (string parentFolder in newNote.ParentFolders)
                if (parentFolder != googleNotesFolder.Self) //Except for Notes root folder
                {
                    Document deletedNote = LoadGoogleNotes(parentFolder + "/contents", newNote.DocumentEntry.Id);
                    //DocumentsRequest.Delete(new Uri(parentFolder + "/contents/" + newNote.ResourceId), newNote.ETag);
                    DocumentsRequest.Delete(deletedNote); //Just delete it from this category
                }

            //Move now to Notes subfolder (if not already there)
            if (!IsInFolder(googleNotesFolder, newNote))
                newNote = DocumentsRequest.MoveDocumentTo(googleNotesFolder, newNote);

            //Move now to all categories subfolder (if not already there)
            foreach (string category in Utilities.GetOutlookGroups(match.OutlookNote.Categories))
            {
                Document categoryFolder = GetOrCreateGoogleFolder(googleNotesFolder, category);

                if (!IsInFolder(categoryFolder, newNote))
                    newNote = DocumentsRequest.MoveDocumentTo(categoryFolder, newNote);
            }

            //Then update the match IDs
            UpdateNoteMatchId(match);

            Logger.Log((create ? "Created" : "Updated") + " Google note from Outlook: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
            //Then release this match as completed (to not log the summary already before each single note result has been synced
            match.AsyncUpdateCompleted = true;
        }

        /// <summary>
        /// returns true, if the passed document is already in passed parentFolder
        /// </summary>
        /// <param name="parentFolder">the parent folder</param>
        /// <param name="childDocument">the document to be checked</param>
        /// <returns></returns>
        private bool IsInFolder(Document parentFolder, Document childDocument)
        {
            foreach (string parent in childDocument.ParentFolders)
            {
                if (parent == parentFolder.Self)
                {
                    return true;
                }
            }
            return false;
        }

        private string GetXml(Contact contact)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                contact.ContactEntry.SaveToXml(ms);
                StreamReader sr = new StreamReader(ms);
                ms.Seek(0, SeekOrigin.Begin);
                return sr.ReadToEnd();
            }
        }

        private static string GetXml(Document note)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                note.DocumentEntry.SaveToXml(ms);
                StreamReader sr = new StreamReader(ms);
                ms.Seek(0, SeekOrigin.Begin);
                string xml = sr.ReadToEnd();
                return xml;
            }
        }

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="googleContact"></param>
        internal Contact SaveGoogleContact(Contact googleContact)
        {
            //check if this contact was not yet inserted on google.
            if (googleContact.ContactEntry.Id.Uri == null)
            {
                //insert contact.
                Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

                try
                {
                    Contact createdEntry = null;

                    try
                    {
                        createdEntry = ContactsRequest.Insert(feedUri, googleContact);
                    }
                    catch (System.Net.ProtocolViolationException)
                    {
                        //TODO (obelix30)
                        //http://stackoverflow.com/questions/23804960/contactsrequest-insertfeeduri-newentry-sometimes-fails-with-system-net-protoc
                        createdEntry = ContactsRequest.Insert(feedUri, googleContact);
                    }

                    return createdEntry;
                }
                catch (GDataRequestException ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    Logger.Log(googleContact, EventType.Debug);
                    string responseString = EscapeXml(ex.ResponseString);
                    string xml = GetXml(googleContact);
                    string newEx = string.Format("Error saving NEW Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    string xml = GetXml(googleContact);
                    string newEx = string.Format("Error saving NEW Google contact:\n{0}\n{1}", ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
            else
            {
                try
                {
                    //contact already present in google. just update

                    // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
                    // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
                    // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
                    int fieldCount = 0;
                    foreach (UserDefinedField userDefinedField in googleContact.ContactEntry.UserDefinedFields)
                    {
                        fieldCount++;
                        if (string.IsNullOrEmpty(userDefinedField.Key))
                        {
                            userDefinedField.Key = "UserField" + fieldCount.ToString();
                            Logger.Log("Set key to user defined field to avoid errors: " + userDefinedField.Key, EventType.Debug);
                        }

                        //similar error with empty values
                        if (string.IsNullOrEmpty(userDefinedField.Value))
                        {
                            userDefinedField.Value = userDefinedField.Key;
                            Logger.Log("Set value to user defined field to avoid errors: " + userDefinedField.Value, EventType.Debug);
                        }
                    }

                    UpdateExtendedProperties(googleContact);

                    //TODO: this will fail if original contact had an empty name or primary email address.

                    Contact updated = null;
                    try
                    {
                        updated = ContactsRequest.Update(googleContact);
                    }
                    catch (System.Net.ProtocolViolationException)
                    {
                        //TODO (obelix30)
                        //http://stackoverflow.com/questions/23804960/contactsrequest-insertfeeduri-newentry-sometimes-fails-with-system-net-protoc
                        updated = ContactsRequest.Update(googleContact);
                    }
                    return updated;
                }
                catch (ApplicationException)
                {
                    throw;
                }
                catch (GDataRequestException ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    Logger.Log(googleContact, EventType.Debug);
                    string responseString = EscapeXml(ex.ResponseString);
                    string xml = GetXml(googleContact);
                    string newEx = string.Format("Error saving EXISTING Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    string xml = GetXml(googleContact);
                    string newEx = string.Format("Error saving EXISTING Google contact:\n{0}\n{1}", ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
        }

        private void UpdateExtendedProperties(Contact googleContact)
        {
            RemoveTooManyExtendedProperties(googleContact);
            RemoveTooBigExtendedProperties(googleContact);
            RemoveDuplicatedExtendedProperties(googleContact);
            UpdateEmptyExtendedProperties(googleContact);
            UpdateTooManyExtendedProperties(googleContact);
            UpdateTooBigExtendedProperties(googleContact);
            UpdateDuplicatedExtendedProperties(googleContact);
        }

        private void UpdateDuplicatedExtendedProperties(Contact googleContact)
        {
            DeleteDuplicatedPropertiesForm form = null;

            try
            {
                HashSet<string> dups = new HashSet<string>();
                foreach (var p in googleContact.ExtendedProperties)
                {
                    if (dups.Contains(p.Name))
                    {
                        Logger.Log(googleContact.Title + ": for extended property " + p.Name + " duplicates were found.", EventType.Debug);
                        if (form == null)
                        {
                            form = new DeleteDuplicatedPropertiesForm();
                        }
                        form.AddExtendedProperty(false, p.Name, "");
                    }
                    else
                    {
                        dups.Add(p.Name);
                    }
                }
                if (form == null)
                    return;

                if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
                {
                    foreach (var p in ContactExtendedPropertiesToRemoveIfDuplicated)
                    {
                        form.AddExtendedProperty(true, p, "");
                    }
                }

                form.SortExtendedProperties();

                if (SettingsForm.Instance.ShowDeleteDuplicatedPropertiesForm(form) == DialogResult.OK)
                {
                    bool allCheck = form.removeFromAll;

                    if (allCheck)
                    {
                        if (ContactExtendedPropertiesToRemoveIfDuplicated == null)
                        {
                            ContactExtendedPropertiesToRemoveIfDuplicated = new HashSet<string>();
                        }
                        else
                        {
                            ContactExtendedPropertiesToRemoveIfDuplicated.Clear();
                        }
                        Logger.Log(googleContact.Title + ": will clean some extended properties for all contacts.", EventType.Debug);
                    }
                    else if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
                    {
                        ContactExtendedPropertiesToRemoveIfDuplicated = null;
                        Logger.Log(googleContact.Title + ": will clean some extended properties for this contact.", EventType.Debug);
                    }

                    foreach (DataGridViewRow r in form.extendedPropertiesRows)
                    {
                        if (Convert.ToBoolean(r.Cells["Selected"].Value))
                        {
                            var key = r.Cells["Key"].Value.ToString();

                            if (allCheck)
                            {
                                ContactExtendedPropertiesToRemoveIfDuplicated.Add(key);
                            }

                            for (var j = googleContact.ExtendedProperties.Count - 1; j >= 0; j--)
                            {
                                if (googleContact.ExtendedProperties[j].Name == key)
                                    googleContact.ExtendedProperties.RemoveAt(j);
                            }

                            Logger.Log("Extended property to remove: " + key, EventType.Debug);
                        }
                    }
                }
            }
            finally
            {
                if (form != null)
                    form.Dispose();
            }
        }

        private void UpdateTooBigExtendedProperties(Contact googleContact)
        {
            DeleteTooBigPropertiesForm form = null;

            try
            {
                foreach (var p in googleContact.ExtendedProperties)
                {
                    if (p.Value.Length > 1012)
                    {
                        Logger.Log(googleContact.Title + ": for extended property " + p.Name + " size limit exceeded (" + p.Value.Length + "). Value is: " + p.Value, EventType.Debug);
                        if (form == null)
                        {
                            form = new DeleteTooBigPropertiesForm();
                        }
                        form.AddExtendedProperty(false, p.Name, p.Value);
                    }
                }
                if (form == null)
                    return;

                if (ContactExtendedPropertiesToRemoveIfTooBig != null)
                {
                    foreach (var p in ContactExtendedPropertiesToRemoveIfTooBig)
                    {
                        form.AddExtendedProperty(true, p, "");
                    }
                }

                form.SortExtendedProperties();

                if (SettingsForm.Instance.ShowDeleteTooBigPropertiesForm(form) == DialogResult.OK)
                {
                    bool allCheck = form.removeFromAll;

                    if (allCheck)
                    {
                        if (ContactExtendedPropertiesToRemoveIfTooBig == null)
                        {
                            ContactExtendedPropertiesToRemoveIfTooBig = new HashSet<string>();
                        }
                        else
                        {
                            ContactExtendedPropertiesToRemoveIfTooBig.Clear();
                        }
                        Logger.Log(googleContact.Title + ": will clean some extended properties for all contacts.", EventType.Debug);
                    }
                    else if (ContactExtendedPropertiesToRemoveIfTooBig != null)
                    {
                        ContactExtendedPropertiesToRemoveIfTooBig = null;
                        Logger.Log(googleContact.Title + ": will clean some extended properties for this contact.", EventType.Debug);
                    }

                    foreach (DataGridViewRow r in form.extendedPropertiesRows)
                    {
                        if (Convert.ToBoolean(r.Cells["Selected"].Value))
                        {
                            var key = r.Cells["Key"].Value.ToString();

                            if (allCheck)
                            {
                                ContactExtendedPropertiesToRemoveIfTooBig.Add(key);
                            }

                            for (var j = googleContact.ExtendedProperties.Count - 1; j >= 0; j--)
                            {
                                if (googleContact.ExtendedProperties[j].Name == key)
                                    googleContact.ExtendedProperties.RemoveAt(j);
                            }

                            Logger.Log("Extended property to remove: " + key, EventType.Debug);
                        }
                    }
                }
            }
            finally
            {
                if (form != null)
                    form.Dispose();
            }
        }

        private void UpdateTooManyExtendedProperties(Contact googleContact)
        {
            if (googleContact.ExtendedProperties.Count > 10)
            {
                Logger.Log(googleContact.Title + ": too many extended properties " + googleContact.ExtendedProperties.Count, EventType.Debug);

                using (DeleteTooManyPropertiesForm form = new DeleteTooManyPropertiesForm())
                {
                    foreach (var p in googleContact.ExtendedProperties)
                    {
                        if (p.Name != "gos:oid:" + SyncProfile)
                            form.AddExtendedProperty(false, p.Name, p.Value);
                    }

                    if (ContactExtendedPropertiesToRemoveIfTooMany != null)
                    {
                        foreach (var p in ContactExtendedPropertiesToRemoveIfTooMany)
                        {
                            form.AddExtendedProperty(true, p, "");
                        }
                    }

                    form.SortExtendedProperties();

                    if (SettingsForm.Instance.ShowDeleteTooManyPropertiesForm(form) == DialogResult.OK)
                    {
                        bool allCheck = form.removeFromAll;

                        if (allCheck)
                        {
                            if (ContactExtendedPropertiesToRemoveIfTooMany == null)
                            {
                                ContactExtendedPropertiesToRemoveIfTooMany = new HashSet<string>();
                            }
                            else
                            {
                                ContactExtendedPropertiesToRemoveIfTooMany.Clear();
                            }
                            Logger.Log(googleContact.Title + ": will clean some extended properties for all contacts.", EventType.Debug);
                        }
                        else if (ContactExtendedPropertiesToRemoveIfTooMany != null)
                        {
                            ContactExtendedPropertiesToRemoveIfTooMany = null;
                            Logger.Log(googleContact.Title + ": will clean some extended properties for this contact.", EventType.Debug);
                        }

                        foreach (DataGridViewRow r in form.extendedPropertiesRows)
                        {
                            if (Convert.ToBoolean(r.Cells["Selected"].Value))
                            {
                                var key = r.Cells["Key"].Value.ToString();

                                if (allCheck)
                                {
                                    ContactExtendedPropertiesToRemoveIfTooMany.Add(key);
                                }

                                for (var i = googleContact.ExtendedProperties.Count - 1; i >= 0; i--)
                                {
                                    if (googleContact.ExtendedProperties[i].Name == key)
                                        googleContact.ExtendedProperties.RemoveAt(i);
                                }

                                Logger.Log("Extended property to remove: " + key, EventType.Debug);
                            }
                        }
                    }
                }
            }
        }

        private static void UpdateEmptyExtendedProperties(Contact googleContact)
        {
            foreach (var p in googleContact.ExtendedProperties)
            {
                if (string.IsNullOrEmpty(p.Value))
                {
                    Logger.Log(googleContact.Title + ": empty value for " + p.Name, EventType.Debug);
                    if (p.ChildNodes != null)
                    {
                        Logger.Log(googleContact.Title + ": childNodes count " + p.ChildNodes.Count, EventType.Debug);
                    }
                    else
                    {
                        p.Value = p.Name;
                        Logger.Log(googleContact.Title + ": set value to extended property to avoid errors " + p.Name, EventType.Debug);
                    }
                }
            }
        }

        private void RemoveDuplicatedExtendedProperties(Contact googleContact)
        {
            if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
            {
                for (var i = googleContact.ExtendedProperties.Count - 1; i >= 0; i--)
                {
                    var key = googleContact.ExtendedProperties[i].Name;
                    if (ContactExtendedPropertiesToRemoveIfDuplicated.Contains(key))
                    {
                        Logger.Log(googleContact.Title + ": removed (duplicate) " + key, EventType.Debug);
                        googleContact.ExtendedProperties.RemoveAt(i);
                    }
                }
            }
        }

        private void RemoveTooBigExtendedProperties(Contact googleContact)
        {
            if (ContactExtendedPropertiesToRemoveIfTooBig != null)
            {
                for (var i = googleContact.ExtendedProperties.Count - 1; i >= 0; i--)
                {
                    if (googleContact.ExtendedProperties[i].Value.Length > 1012)
                    {
                        var key = googleContact.ExtendedProperties[i].Name;
                        if (ContactExtendedPropertiesToRemoveIfTooBig.Contains(key))
                        {
                            Logger.Log(googleContact.Title + ": removed (size)" + key, EventType.Debug);
                            googleContact.ExtendedProperties.RemoveAt(i);
                        }
                    }
                }
            }
        }

        private void RemoveTooManyExtendedProperties(Contact googleContact)
        {
            if (ContactExtendedPropertiesToRemoveIfTooMany != null)
            {
                for (var i = googleContact.ExtendedProperties.Count - 1; i >= 0; i--)
                {
                    var key = googleContact.ExtendedProperties[i].Name;
                    if (ContactExtendedPropertiesToRemoveIfTooMany.Contains(key))
                    {
                        Logger.Log(googleContact.Title + ": removed (count) " + key, EventType.Debug);
                        googleContact.ExtendedProperties.RemoveAt(i);
                    }
                }
            }
        }

        /// <summary>
        /// Save the google Appointment
        /// </summary>
        /// <param name="googleAppointment"></param>
        internal Google.Apis.Calendar.v3.Data.Event SaveGoogleAppointment(Google.Apis.Calendar.v3.Data.Event googleAppointment)
        {
            //check if this contact was not yet inserted on google.
            if (googleAppointment.Id == null)
            {
                ////insert contact.
                //Uri feedUri = new Uri("https://www.google.com/calendar/feeds/default/private/full");

                try
                {
                    Google.Apis.Calendar.v3.Data.Event createdEntry = EventRequest.Insert(googleAppointment, SyncAppointmentsGoogleFolder).Execute();
                    return createdEntry;
                }
                catch (Exception ex)
                {
                    Logger.Log(googleAppointment, EventType.Debug);
                    string newEx = string.Format("Error saving NEW Google appointment: {0}. \n{1}", googleAppointment.Summary + " - " + GetTime(googleAppointment), ex.Message);
                    throw new ApplicationException(newEx, ex);
                }
            }
            else
            {
                try
                {
                    //contact already present in google. just update
                    Google.Apis.Calendar.v3.Data.Event updated = EventRequest.Update(googleAppointment, SyncAppointmentsGoogleFolder, googleAppointment.Id).Execute();
                    return updated;
                }
                catch (Exception ex)
                {
                    Logger.Log(googleAppointment, EventType.Debug);

                    string error = "Error saving EXISTING Google appointment: ";
                    error += googleAppointment.Summary + " - " + GetTime(googleAppointment);
                    error += " - Creator: " + (googleAppointment.Creator != null ? googleAppointment.Creator.Email : "null");
                    error += " - Organizer: " + (googleAppointment.Organizer != null ? googleAppointment.Organizer.Email : "null");
                    error += ". \n" + ex.Message;
                    Logger.Log(error, EventType.Warning);
                    //string newEx = String.Format("Error saving EXISTING Google appointment: {0}. \n{1}", googleAppointment.Summary + " - " + GetTime(googleAppointment), ex.Message);
                    //throw new ApplicationException(newEx, ex);

                    return googleAppointment;
                }
            }
        }

        /// <summary>
        /// save the google note
        /// </summary>
        /// <param name="googleNote"></param>
        public static Document SaveGoogleNote(Document parentFolder, Document googleNote, DocumentsRequest documentsRequest)
        {
            //check if this contact was not yet inserted on google.
            if (googleNote.DocumentEntry.Id.Uri == null)
            {
                //insert contact.
                Uri feedUri = null;

                if (parentFolder != null)
                {
                    try
                    {//In case of Notes folder creation, the GoogleNotesFolder.DocumentEntry.Content.AbsoluteUri throws a NullReferenceException
                        feedUri = new Uri(parentFolder.DocumentEntry.Content.AbsoluteUri);
                    }
                    catch (Exception)
                    { }
                }

                if (feedUri == null)
                    feedUri = new Uri(documentsRequest.BaseUri);

                try
                {
                    Document createdEntry = documentsRequest.Insert(feedUri, googleNote);
                    //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, createdEntry, googleNote);    
                    Logger.Log("Created new Google folder: " + createdEntry.Title, EventType.Information);
                    return createdEntry;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    GDataRequestException e = ex as GDataRequestException;
                    if (e != null)
                        responseString = EscapeXml(e.ResponseString);
                    string xml = GetXml(googleNote);
                    string newEx = string.Format("Error saving NEW Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
            else
            {
                try
                {
                    //note already present in google. just update
                    Document updated = documentsRequest.Update(googleNote);

                    //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, updated, googleNote);                   

                    return updated;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    GDataRequestException e = ex as GDataRequestException;
                    if (e != null)
                        responseString = EscapeXml(e.ResponseString);
                    string xml = GetXml(googleNote);
                    string newEx = string.Format("Error saving EXISTING Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
        }

        //public void SaveContactPhotos(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (!hasGooglePhoto && !hasOutlookPhoto)
        //        return;
        //    else if (hasGooglePhoto && _syncOption != SyncOption.OutlookToGoogleOnly)
        //    {
        //        // add google photo to outlook
        //        Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
        //        match.OutlookContact.Save();

        //        googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto && _syncOption != SyncOption.GoogleToOutlookOnly)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else
        //    {
        //        // TODO: if both contacts have photos and one is updated, the
        //        // other will not be updated.
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveGooglePhoto(ContactMatch match, Outlook.ContactItem outlookContactitem)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(outlookContactitem);

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                using (var outlookPhoto = Utilities.GetOutlookPhoto(outlookContactitem))
                {
                    if (outlookPhoto != null)
                    {
                        //Try up to 5 times to overcome Google issue
                        for (int retry = 0; retry < 5; retry++)
                        {
                            try
                            {
                                using (var bmp = new Bitmap(outlookPhoto))
                                {
                                    using (var stream = new MemoryStream(Utilities.BitmapToBytes(bmp)))
                                    {
                                        // Save image to stream.
                                        //outlookPhoto.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);

                                        //Don'T crop, because maybe someone wants to keep his photo like it is on Outlook
                                        //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);                        
                                        ContactsRequest.SetPhoto(match.GoogleContact, stream);

                                        //Just save the Outlook Contact to have the same lastUpdate date as Google
                                        ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactitem, match.GoogleContact);
                                        outlookContactitem.Save();
                                        outlookPhoto.Dispose();

                                    }
                                }
                                break; //Exit because photo save succeeded
                            }
                            catch (GDataRequestException ex)
                            { //If Google found a picture for a new Google account, it sets it automatically and throws an error, if updating it with the Outlook photo. 
                              //Therefore save it again and try again to save the photo
                                if (retry == 4)
                                    ErrorHandler.Handle(new Exception("Photo of contact " + match.GoogleContact.Title + "couldn't be saved after 5 tries, maybe Google found its own photo and doesn't allow updating it", ex));
                                else
                                {
                                    Thread.Sleep(1000);
                                    //LoadGoogleContact again to get latest ETag
                                    //match.GoogleContact = LoadGoogleContacts(match.GoogleContact.AtomEntry.Id);
                                    match.GoogleContact = SaveGoogleContact(match.GoogleContact);
                                }
                            }
                        }
                    }
                }
            }
            else if (hasGooglePhoto)
            {
                //Delete Photo on Google side, if no Outlook photo exists
                ContactsRequest.Delete(match.GoogleContact.PhotoUri, match.GoogleContact.PhotoEtag);
            }

            Utilities.DeleteTempPhoto();
        }

        //public void SaveOutlookPhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasGooglePhoto)
        //    {
        //        Image image = new Image(match.GoogleContact.PhotoUri);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, image);
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();

        //        //googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto)
        //    {
        //        match.OutlookContact.RemovePicture();
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();
        //    }
        //}

        //public void SaveGooglePhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasOutlookPhoto)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            //Just save the Outlook Contact to have the same lastUpdate date as Google
        //            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //            match.OutlookContact.Save();
        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else if (hasGooglePhoto)
        //    {
        //        //ToDo: Delete Photo on Google side, if no Outlook photo exists
        //        //match.GoogleContact.PhotoUri = null;
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveOutlookPhoto(Contact googleContact, Outlook.ContactItem outlookContact)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(googleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(outlookContact);

            if (hasGooglePhoto)
            {
                // add google photo to outlook
                //ToDo: add google photo to outlook with new Google API
                //Stream stream = _googleService.GetPhoto(match.GoogleContact);
                using (var googlePhoto = Utilities.GetGooglePhoto(this, googleContact))
                {
                    if (googlePhoto != null)    // Google may have an invalid photo
                    {
                        Utilities.SetOutlookPhoto(outlookContact, googlePhoto);
                        ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
                        outlookContact.Save();
                    }
                }
            }
            else if (hasOutlookPhoto)
            {
                outlookContact.RemovePicture();
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
                outlookContact.Save();
            }
        }

        public Group SaveGoogleGroup(Group group)
        {
            //check if this group was not yet inserted on google.
            if (group.GroupEntry.Id.Uri == null)
            {
                //insert group.
                Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

                try
                {
                    return ContactsRequest.Insert(feedUri, group);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    Logger.Log("Group dump: " + group.ToString(), EventType.Debug);
                    throw;
                }
            }
            else
            {
                try
                {
                    //group already present in google. just update
                    return ContactsRequest.Update(group);
                }
                catch
                {
                    //TODO: save google group xml for diagnistics
                    throw;
                }
            }
        }

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        public void UpdateContact(Outlook.ContactItem master, Contact slave)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);
        }

        /// <summary>
        /// Updates Outlook contact from Google (including groups/categories)
        /// </summary>
        public void UpdateContact(Contact master, Outlook.ContactItem slave)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);

            // -- Immediately save the Outlook contact (including groups) so it can be released, and don't do it in the save loop later
            SaveOutlookContact(ref master, slave);
            SyncedCount++;
            Logger.Log("Updated Outlook contact from Google: \"" + slave.FileAs + "\".", EventType.Information);
        }

        /// <summary>
        /// Updates Google note from Outlook
        /// </summary>
        public void UpdateNote(Outlook.NoteItem master, Document slave)
        {
            if (!string.IsNullOrEmpty(master.Subject))
                slave.Title = master.Subject.Replace(":", string.Empty);

            string fileName = NotePropertiesUtils.CreateNoteFile(master.EntryID, master.Body, SyncProfile);

            string contentType = MediaFileSource.GetContentTypeForFileName(fileName);

            //ToDo: Somewhow, the content is not uploaded to Google, only an empty document
            //Therefoe I use DocumentService.UploadDocument instead.
            slave.MediaSource = new MediaFileSource(fileName, contentType);
        }

        /// <summary>
        /// Updates Outlook contact from Google
        /// </summary>
        public void UpdateNote(Document master, Outlook.NoteItem slave)
        {
            //slave.Subject = master.Title; //The Subject is readonly and set automatically by Outlook
            string body = NotePropertiesUtils.GetBody(this, master);

            if (string.IsNullOrEmpty(body) && slave.Body != null)
            {
                //DialogResult result = MessageBox.Show("The body of Google note '" + master.Title + "' is empty. Do you really want to synchronize an empty Google note to a not yet empty Outlook note?", "Empty Google Note", MessageBoxButtons.YesNo);

                //if (result != DialogResult.Yes)
                //{
                //    Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to skip this note and not to synchronize an empty Google note to a not yet empty Outlook note.", EventType.Information);
                Logger.Log("The body of Google note '" + master.Title + "' is empty. It is skipped from syncing, because Outlook note is not empty.", EventType.Warning);
                SkippedCount++;
                return;
                //}
                //Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to synchronize an empty Google note to a not yet empty Outlook note (" + slave.Body + ").", EventType.Warning);                

            }

            slave.Body = body;

            slave.Categories = string.Empty;
            List<string> newCats = new List<string>();
            foreach (string category in master.ParentFolders)
            {
                Document categoryFolder = GetGoogleFolder(googleNotesFolder, null, category);

                if (categoryFolder != null)
                    newCats.Add(categoryFolder.Title);

            }

            slave.Categories = string.Join(", ", newCats.ToArray());

            NotePropertiesUtils.CreateNoteFile(master.Id, body, SyncProfile);
        }

        /// <summary>
        /// Updates Google contact's groups from Outlook contact
        /// </summary>
        private void OverwriteContactGroups(Outlook.ContactItem master, Contact slave)
        {
            Collection<Group> currentGroups = Utilities.GetGoogleGroups(this, slave);

            // get outlook categories
            string[] cats = Utilities.GetOutlookGroups(master.Categories);

            // remove obsolete groups
            Collection<Group> remove = new Collection<Group>();
            bool found;
            foreach (Group group in currentGroups)
            {
                found = false;
                foreach (string cat in cats)
                {
                    if (group.Title == cat)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    remove.Add(group);
            }
            while (remove.Count != 0)
            {
                Utilities.RemoveGoogleGroup(slave, remove[0]);
                remove.RemoveAt(0);
            }

            // add new groups
            Group g;
            foreach (string cat in cats)
            {
                if (!Utilities.ContainsGroup(this, slave, cat))
                {
                    // add group to contact
                    g = GetGoogleGroupByName(cat);
                    if (g == null)
                    {
                        // try to create group again (if not yet created before
                        g = CreateGroup(cat);

                        if (g != null)
                        {
                            g = SaveGoogleGroup(g);
                            if (g != null)
                                GoogleGroups.Add(g);
                            else
                                Logger.Log("Google Groups were supposed to be created prior to saving a contact. Unfortunately the group '" + cat + "' couldn't be saved on Google side and was not assigned to the contact: " + master.FileAs, EventType.Warning);
                        }
                        else
                            Logger.Log("Google Groups were supposed to be created prior to saving a contact. Unfortunately the group '" + cat + "' couldn't be created and was not assigned to the contact: " + master.FileAs, EventType.Warning);

                    }

                    if (g != null)
                        Utilities.AddGoogleGroup(slave, g);
                }
            }

            //add system Group My Contacts            
            if (!Utilities.ContainsGroup(this, slave, myContactsGroup))
            {
                // add group to contact
                g = GetGoogleGroupByName(myContactsGroup);
                if (g == null)
                {
                    throw new Exception(string.Format("Google {0} doesn't exist", myContactsGroup));
                }
                Utilities.AddGoogleGroup(slave, g);
            }
        }

        /// <summary>
        /// Updates Outlook contact's categories (groups) from Google groups
        /// </summary>
        private void OverwriteContactGroups(Contact master, Outlook.ContactItem slave)
        {
            Collection<Group> newGroups = Utilities.GetGoogleGroups(this, master);

            List<string> newCats = new List<string>(newGroups.Count);
            foreach (Group group in newGroups)
            {   //Only add groups that are no SystemGroup (e.g. "System Group: Meine Kontakte") automatically tracked by Google
                if (group.Title != null && !group.Title.Equals(myContactsGroup))
                    newCats.Add(group.Title);
            }

            slave.Categories = string.Join(", ", newCats.ToArray());
        }

        /// <summary>
        /// Resets associantions of Outlook contacts with Google contacts via user props
        /// and resets associantions of Google contacts with Outlook contacts via extended properties.
        /// </summary>
        public void ResetContactMatches()
        {
            Debug.Assert(OutlookContacts != null, "Outlook Contacts object is null - this should not happen. Please inform Developers.");
            Debug.Assert(GoogleContacts != null, "Google Contacts object is null - this should not happen. Please inform Developers.");

            try
            {
                if (string.IsNullOrEmpty(SyncProfile))
                {
                    Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                    return;
                }

                lock (_syncRoot)
                {
                    Logger.Log("Resetting Google Contact matches...", EventType.Information);
                    foreach (Contact googleContact in GoogleContacts)
                    {
                        try
                        {
                            if (googleContact != null)
                                ResetMatch(googleContact);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("The match of Google contact " + ContactMatch.GetName(googleContact) + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
                    }

                    Logger.Log("Resetting Outlook Contact matches...", EventType.Information);
                    //1 based array
                    for (int i = 1; i <= OutlookContacts.Count; i++)
                    {
                        Outlook.ContactItem outlookContact = null;

                        try
                        {
                            outlookContact = OutlookContacts[i] as Outlook.ContactItem;
                            if (outlookContact == null)
                            {
                                Logger.Log("Empty Outlook contact found (maybe distribution list). Skipping", EventType.Warning);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            //this is needed because some contacts throw exceptions
                            Logger.Log("Accessing Outlook contact threw and exception. Skipping: " + ex.Message, EventType.Warning);
                            continue;
                        }

                        try
                        {
                            ResetMatch(outlookContact);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("The match of Outlook contact " + outlookContact.FileAs + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
                    }

                }
            }
            finally
            {
                if (OutlookContacts != null)
                {
                    Marshal.ReleaseComObject(OutlookContacts);
                    OutlookContacts = null;
                }
                GoogleContacts = null;
            }

        }

        /// <summary>
        /// Resets associantions of Outlook notes with Google contacts via user props
        /// and resets associantions of Google contacts with Outlook contacts via extended properties.
        /// </summary>
        public void ResetNoteMatches()
        {
            Debug.Assert(OutlookNotes != null, "Outlook Notes object is null - this should not happen. Please inform Developers.");

            //try
            //{
            if (string.IsNullOrEmpty(SyncProfile))
            {
                Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                return;
            }


            lock (_syncRoot)
            {
                Logger.Log("Resetting Google Note matches...", EventType.Information);

                try
                {
                    NotePropertiesUtils.DeleteNoteFiles(SyncProfile);
                }
                catch (Exception ex)
                {
                    Logger.Log("The Google Note matches couldn't be reset: " + ex.Message, EventType.Warning);
                }


                Logger.Log("Resetting Outlook Note matches...", EventType.Information);
                //1 based array
                for (int i = 1; i <= OutlookNotes.Count; i++)
                {
                    Outlook.NoteItem outlookNote = null;

                    try
                    {
                        outlookNote = OutlookNotes[i] as Outlook.NoteItem;
                        if (outlookNote == null)
                        {
                            Logger.Log("Empty Outlook Note found (maybe distribution list). Skipping", EventType.Warning);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some notes throw exceptions
                        Logger.Log("Accessing Outlook Note threw and exception. Skipping: " + ex.Message, EventType.Warning);
                        continue;
                    }

                    try
                    {
                        ResetMatch(outlookNote);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log("The match of Outlook note " + outlookNote.Subject + " couldn't be reset: " + ex.Message, EventType.Warning);
                    }
                }

            }
            //}
            //finally
            //{
            //    if (OutlookContacts != null)
            //    {
            //        Marshal.ReleaseComObject(OutlookContacts);
            //        OutlookContacts = null;
            //    }
            //    GoogleContacts = null;
            //}

        }


        ///// <summary>
        ///// Reset the match link between Google and Outlook contact
        ///// </summary>
        ///// <param name="match"></param>
        //public void ResetMatch(ContactMatch match)
        //{           
        //    if (match == null)
        //        throw new ArgumentNullException("match", "Given ContactMatch is null");


        //    if (match.GoogleContact != null)
        //    {
        //        ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, match.GoogleContact);
        //        SaveGoogleContact(match.GoogleContact);
        //    }

        //    if (match.OutlookContact != null)
        //    {
        //        Outlook.ContactItem outlookContactItem = match.OutlookContact.GetOriginalItemFromOutlook(this);
        //        try
        //        {
        //            ContactPropertiesUtils.ResetOutlookGoogleContactId(this, outlookContactItem);
        //            outlookContactItem.Save();
        //        }
        //        finally
        //        {
        //            Marshal.ReleaseComObject(outlookContactItem);
        //            outlookContactItem = null;
        //        }

        //        //Reset also Google duplicatesC
        //        foreach (Contact duplicate in match.AllGoogleContactMatches)
        //        {
        //            if (duplicate != match.GoogleContact)
        //            { //To save time, only if not match.GoogleContact, because this was already reset above
        //                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, duplicate);
        //                SaveGoogleContact(duplicate);
        //            }
        //        }
        //    }


        //}

        /// <summary>
        /// Resets associations of Outlook appointments with Google appointments via user props
        /// and vice versa
        /// </summary>
        public void ResetOutlookAppointmentMatches(bool deleteOutlookAppointments)
        {
            Debug.Assert(OutlookAppointments != null, "Outlook Appointments object is null - this should not happen. Please inform Developers.");

            //try
            //{

            lock (_syncRoot)
            {

                Logger.Log("Resetting Outlook appointment matches...", EventType.Information);
                //1 based array
                for (int i = OutlookAppointments.Count; i >= 1; i--)
                {
                    Outlook.AppointmentItem outlookAppointment = null;

                    try
                    {
                        outlookAppointment = OutlookAppointments[i] as Outlook.AppointmentItem;
                        if (outlookAppointment == null)
                        {
                            Logger.Log("Empty Outlook appointment found (maybe distribution list). Skipping", EventType.Warning);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some appointments throw exceptions
                        Logger.Log("Accessing Outlook appointment threw an exception. Skipping: " + ex.Message, EventType.Warning);
                        continue;
                    }

                    if (deleteOutlookAppointments)
                    {
                        outlookAppointment.Delete();
                    }
                    else
                    {
                        try
                        {
                            ResetMatch(outlookAppointment);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("The match of Outlook appointment " + outlookAppointment.Subject + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
                    }
                }

            }
            //}
            //finally
            //{
            //    if (OutlookContacts != null)
            //    {
            //        Marshal.ReleaseComObject(OutlookContacts);
            //        OutlookContacts = null;
            //    }
            //    GoogleContacts = null;
            //}

        }

        /// <summary>
        /// Reset the match link between Google and Outlook contact        
        /// </summary>
        public Contact ResetMatch(Contact googleContact)
        {

            if (googleContact != null)
            {
                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, googleContact);
                return SaveGoogleContact(googleContact);
            }
            else
                return googleContact;
        }

        public Google.Apis.Calendar.v3.Data.Event ResetMatch(Google.Apis.Calendar.v3.Data.Event googleAppointment)
        {

            if (googleAppointment != null)
            {
                AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, googleAppointment);
                return SaveGoogleAppointment(googleAppointment);
            }
            else
                return googleAppointment;
        }

        /// <summary>
        /// Reset the match link between Outlook and Google contact
        /// </summary>
        public void ResetMatch(Outlook.ContactItem outlookContact)
        {
            if (outlookContact != null)
            {
                try
                {
                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, outlookContact);
                    outlookContact.Save();
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookContact);
                    outlookContact = null;
                }
            }
        }

        /// <summary>
        /// Reset the match link between Outlook and Google note
        /// </summary>
        public void ResetMatch(Outlook.NoteItem outlookNote)
        {

            if (outlookNote != null)
            {
                //try
                //{
                NotePropertiesUtils.ResetOutlookGoogleNoteId(this, outlookNote);
                outlookNote.Save();
                //}
                //finally
                //{
                //    Marshal.ReleaseComObject(outlookNote);
                //    outlookNote = null;
                //}
            }
        }

        /// <summary>
        /// Reset the match link between Outlook and Google appointment
        /// </summary>
        public void ResetMatch(Outlook.AppointmentItem outlookAppointment)
        {
            if (outlookAppointment != null)
            {
                //try
                //{
                AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, outlookAppointment);
                outlookAppointment.Save();
                //}
                //finally
                //{
                //    Marshal.ReleaseComObject(OutlookAppointment);
                //    OutlookAppointment = null;
                //}
            }
        }

        public ContactMatch ContactByProperty(string name, string email)
        {
            foreach (ContactMatch m in Contacts)
            {
                if (m.GoogleContact != null &&
                    ((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
                    m.GoogleContact.Title == name ||
                    m.GoogleContact.Name != null && m.GoogleContact.Name.FullName == name))
                {
                    return m;
                }
                else if (m.OutlookContact != null && (
                    (m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email) ||
                    m.OutlookContact.FileAs == name))
                {
                    return m;
                }
            }
            return null;
        }

        //public ContactMatch ContactEmail(string email)
        //{
        //    foreach (ContactMatch m in Contacts)
        //    {
        //        if (m.GoogleContact != null &&
        //            (m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email))
        //        {
        //            return m;
        //        }
        //        else if (m.OutlookContact != null && (
        //            m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email))
        //        {
        //            return m;
        //        }
        //    }
        //    return null;
        //}

        /// <summary>
        /// Used to find duplicates.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public Collection<OutlookContactInfo> OutlookContactByProperty(string name, string value)
        {
            Collection<OutlookContactInfo> col = new Collection<OutlookContactInfo>();
            //foreach (Outlook.ContactItem outlookContact in OutlookContacts)
            //{
            //    if (outlookContact != null && (
            //        (outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
            //        outlookContact.FileAs == name))
            //    {
            //        col.Add(outlookContact);
            //    }
            //}
            Outlook.ContactItem item = null;
            try
            {
                item = OutlookContacts.Find("[" + name + "] = \"" + value + "\"") as Outlook.ContactItem;
                while (item != null)
                {
                    col.Add(new OutlookContactInfo(item, this));
                    Marshal.ReleaseComObject(item);
                    item = OutlookContacts.FindNext() as Outlook.ContactItem;
                }
            }
            catch (Exception)
            {
                //TODO: should not get here.
            }

            return col;
        }

        public Group GetGoogleGroupById(string id)
        {
            //return GoogleGroups.FindById(new AtomId(id)) as Group;
            AtomId atomId = new AtomId(id);
            foreach (Group group in GoogleGroups)
            {
                if (group.GroupEntry.Id.Equals(atomId))
                    return group;
            }
            return null;
        }

        public Group GetGoogleGroupByName(string name)
        {
            foreach (Group group in GoogleGroups)
            {
                if (group.Title == name)
                    return group;
            }
            return null;
        }

        public Contact GetGoogleContactById(string id)
        {
            AtomId atomId = new AtomId(id);
            foreach (Contact contact in GoogleContacts)
            {
                if (contact.ContactEntry.Id.Equals(atomId))
                    return contact;
            }
            return null;
        }

        public Document GetGoogleNoteById(string id)
        {
            AtomId atomId = new AtomId(id);
            foreach (Document note in GoogleNotes)
            {
                if (note.DocumentEntry.Id.Equals(atomId))
                    return note;
            }
            return null;
        }

        public Google.Apis.Calendar.v3.Data.Event GetGoogleAppointmentById(string id)
        {
            //ToDo: Temporary remove prefix used by v2:
            id = id.Replace("http://www.google.com/calendar/feeds/default/events/", "");
            id = id.Replace("https://www.google.com/calendar/feeds/default/events/", "");

            //AtomId atomId = new AtomId(id);
            foreach (Google.Apis.Calendar.v3.Data.Event appointment in GoogleAppointments)
            {
                if (appointment.Id.Equals(id))
                    return appointment;
            }

            if (AllGoogleAppointments != null)
                foreach (Google.Apis.Calendar.v3.Data.Event appointment in AllGoogleAppointments)
                {
                    if (appointment.Id.Equals(id))
                        return appointment;
                }

            return null;
        }

        public Outlook.AppointmentItem GetOutlookAppointmentById(string id)
        {
            for (int i = OutlookAppointments.Count; i >= 1; i--)
            {
                Outlook.AppointmentItem a = null;

                try
                {
                    a = OutlookAppointments[i] as Outlook.AppointmentItem;
                    if (a == null)
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                if (AppointmentPropertiesUtils.GetOutlookId(a) == id)
                    return a;
            }
            return null;
        }

        public Outlook.ContactItem GetOutlookContactById(string id)
        {
            for (int i = OutlookContacts.Count; i >= 1; i--)
            {
                Outlook.ContactItem a = null;

                try
                {
                    a = OutlookContacts[i] as Outlook.ContactItem;
                    if (a == null)
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                if (ContactPropertiesUtils.GetOutlookId(a) == id)
                    return a;
            }
            return null;
        }

        //public Event GetGoogleAppointmentByStartDate(AtomId id, DateTime restrictStartDate)
        //{//ToDo: Doesn't work for all recurrences

        //    if (id == null)
        //        return null;

        //    foreach (Event appointment in GoogleAppointments)
        //    {

        //        if (appointment.OriginalEvent != null && appointment.Times.Count > 0 && restrictStartDate.Date.Equals(appointment.Times[0].StartTime.Date))
        //            if (id.Equals(new AtomId(id.AbsoluteUri.Substring(0, id.AbsoluteUri.LastIndexOf("/") + 1) + appointment.OriginalEvent.IdOriginal)))
        //                return appointment;                                         
        //    }

        //    //If not found, load AllGoogleAppointments
        //    if (AllGoogleAppointments == null)
        //        LoadGoogleAppointments(null, 0, 0, null, null);
        //    foreach (Event appointment in AllGoogleAppointments)
        //    {
        //        if (appointment.OriginalEvent != null && appointment.Times.Count > 0 && restrictStartDate.Date.Equals(appointment.Times[0].StartTime.Date))
        //            if (id.Equals(new AtomId(id.AbsoluteUri.Substring(0, id.AbsoluteUri.LastIndexOf("/") + 1) + appointment.OriginalEvent.IdOriginal)))
        //                return appointment;      
        //    }

        //    return null;
        //}

        public Group CreateGroup(string name)
        {
            Group group = new Group();
            group.Title = name;
            group.GroupEntry.Dirty = true;
            return group;
        }

        public static bool AreEqual(Outlook.ContactItem c1, Outlook.ContactItem c2)
        {
            return c1.Email1Address == c2.Email1Address;
        }

        public static int IndexOf(Collection<Outlook.ContactItem> col, Outlook.ContactItem outlookContact)
        {
            for (int i = 0; i < col.Count; i++)
            {
                if (AreEqual(col[i], outlookContact))
                    return i;
            }
            return -1;
        }

        internal void DebugContacts()
        {
            string msg = "DEBUG INFORMATION\nPlease submit to developer:\n\n{0}\n{1}\n{2}";

            if (SyncContacts)
            {
                string oCount = "Outlook Contact Count: " + OutlookContacts.Count;
                string gCount = "Google Contact Count: " + GoogleContacts.Count;
                string mCount = "Matches Count: " + Contacts.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (SyncNotes)
            {
                string oCount = "Outlook Notes Count: " + OutlookNotes.Count;
                string gCount = "Google Notes Count: " + GoogleNotes.Count;
                string mCount = "Matches Count: " + Notes.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (SyncAppointments)
            {
                string oCount = "Outlook appointments Count: " + OutlookAppointments.Count;
                string gCount = "Google appointments Count: " + GoogleAppointments.Count;
                string mCount = "Matches Count: " + Appointments.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public static Outlook.ContactItem CreateOutlookContactItem(string syncContactsFolder)
        {
            //outlookContact = OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.ContactItem outlookContact = null;
            Outlook.MAPIFolder contactsFolder = null;
            Outlook.Items items = null;

            try
            {
                contactsFolder = OutlookNameSpace.GetFolderFromID(syncContactsFolder);
                items = contactsFolder.Items;
                outlookContact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
            }
            return outlookContact;
        }

        public static Outlook.NoteItem CreateOutlookNoteItem(string syncNotesFolder)
        {
            //outlookNote = OutlookApplication.CreateItem(Outlook.OlItemType.olNoteItem) as Outlook.NoteItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.NoteItem outlookNote = null;
            Outlook.MAPIFolder notesFolder = null;
            Outlook.Items items = null;

            try
            {
                notesFolder = OutlookNameSpace.GetFolderFromID(syncNotesFolder);
                items = notesFolder.Items;
                outlookNote = items.Add(Outlook.OlItemType.olNoteItem) as Outlook.NoteItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (notesFolder != null) Marshal.ReleaseComObject(notesFolder);
            }
            return outlookNote;
        }


        public static Outlook.AppointmentItem CreateOutlookAppointmentItem(string syncAppointmentsFolder)
        {
            //OutlookAppointment = OutlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.AppointmentItem outlookAppointment = null;
            Outlook.MAPIFolder appointmentsFolder = null;
            Outlook.Items items = null;

            try
            {
                appointmentsFolder = OutlookNameSpace.GetFolderFromID(syncAppointmentsFolder);
                items = appointmentsFolder.Items;
                outlookAppointment = items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (appointmentsFolder != null) Marshal.ReleaseComObject(appointmentsFolder);
            }
            return outlookAppointment;
        }

        public void Dispose()
        {
            ((IDisposable)CalendarRequest).Dispose();
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
