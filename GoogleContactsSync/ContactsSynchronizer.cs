using Google.Contacts;
using Google.GData.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using Google.GData.Contacts;
using System.IO;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    class ContactsSynchronizer
    {
        internal const string myContactsGroup = "System Group: My Contacts";
        public delegate void DuplicatesFoundHandler(string title, string message);
        public event DuplicatesFoundHandler DuplicatesFound;
        public event ErrorNotificationHandler ErrorEncountered;
        public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);

        public static bool SyncContactsForceRTF { get; set; }

        /// <summary>
        /// if true, use Outlook's FileAs for Google Title/FullName. If false, use Outlook's Fullname
        /// </summary>
        public bool UseFileAs { get; set; }

        public int ErrorCount { get; set; }
        public int SyncedCount { get; set; }
        public int SkippedCount { get; set; }
        public int SkippedCountNotMatches { get; set; }

        private SyncOption _syncOption = SyncOption.MergeOutlookWins;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; }
        }
        public string SyncProfile { get; set; }

        public string OutlookPropertyNameId
        {
            get { return OutlookPropertyPrefix + "id"; }
        }
        public string OutlookPropertyPrefix { get; set; }
        public string OutlookPropertyNameSynced
        {
            get { return OutlookPropertyPrefix + "up"; }
        }

        public ConflictResolution ConflictResolution { get; set; }
        public DeleteResolution DeleteGoogleResolution { get; set; }
        public DeleteResolution DeleteOutlookResolution { get; set; }

        public ContactsRequest ContactsRequest { get; set; }
        public Outlook.Items OutlookContacts { get; set; }

        public Collection<ContactMatch> OutlookContactDuplicates { get; set; }
        public Collection<ContactMatch> GoogleContactDuplicates { get; set; }
        public Collection<Contact> GoogleContacts { get; set; }

        public Collection<Group> GoogleGroups { get; set; }

        public static string SyncContactsFolder { get; set; }

        public List<ContactMatch> Contacts { get; set; }

        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooMany = null;
        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooBig = null;
        private HashSet<string> ContactExtendedPropertiesToRemoveIfDuplicated = null;

        public int DeletedCount { get; set; }
        public bool SyncDelete { get; set; }
        public bool PromptDelete { get; set; }
        public int TotalCount { get; set; }

        private void LoadOutlookContacts()
        {
            Logger.Log("Loading Outlook contacts...", EventType.Information);
            OutlookContacts = Synchronizer.GetOutlookItems(Outlook.OlDefaultFolders.olFolderContacts, SyncContactsFolder);
            Logger.Log("Outlook Contacts Found: " + OutlookContacts.Count, EventType.Debug);
        }

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

        internal void LoadGoogleGroups()
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

        /// <summary>
        /// Load the contacts from Google and Outlook
        /// </summary>
        public void LoadContacts()
        {
            LoadOutlookContacts();
            LoadGoogleGroups();
            LoadGoogleContacts();
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

        public void ResolveDuplicateContacts(Collection<ContactMatch> googleContactDuplicates)
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

        public void SaveGoogleContact(ContactMatch match)
        {
            Outlook.ContactItem olc = null;
            try
            {
                olc = match.OutlookContact.GetOriginalItemFromOutlook();
                ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, olc);
                match.GoogleContact = SaveGoogleContact(match.GoogleContact);
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, olc, match.GoogleContact);
                olc.Save();

                //Now save the Photo
                SaveGooglePhoto(match, olc);
            }
            finally
            {
                if (olc != null)
                {
                    Marshal.ReleaseComObject(olc);
                    olc = null;
                }
            }
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

                    UpdateEmptyUserProperties(googleContact);

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
                    if (p.Value != null && p.Value.Length > 1012)
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

        private static void UpdateEmptyUserProperties(Contact googleContact)
        {
            // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
            // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
            // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
            if (googleContact.ContactEntry == null)
                return;

            if (googleContact.ContactEntry.UserDefinedFields == null)
                return;

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

        public void SaveGooglePhoto(ContactMatch match, Outlook.ContactItem olc)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(olc);

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                using (var outlookPhoto = Utilities.GetOutlookPhoto(olc))
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
                                        ContactPropertiesUtils.SetOutlookGoogleContactId(this, olc, match.GoogleContact);
                                        olc.Save();
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
        }

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
                catch (Exception ex)
                {
                    Logger.Log(ex, EventType.Debug);
                    Logger.Log("Group dump: " + group.ToString(), EventType.Debug);
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

                lock (Synchronizer._syncRoot)
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

        /// <summary>
        /// Used to find duplicates.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public Collection<OutlookContactInfo> OutlookContactByProperty(string name, string value)
        {
            Collection<OutlookContactInfo> col = new Collection<OutlookContactInfo>();

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

        public Group CreateGroup(string name)
        {
            var group = new Group();
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

        public static Outlook.ContactItem CreateOutlookContactItem(string syncContactsFolder)
        {
            //outlookContact = OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.ContactItem outlookContact = null;
            Outlook.MAPIFolder contactsFolder = null;
            Outlook.Items items = null;

            try
            {
                contactsFolder = Synchronizer.OutlookNameSpace.GetFolderFromID(syncContactsFolder);
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

        private static string EscapeXml(string xml)
        {
            return System.Security.SecurityElement.Escape(xml);
        }
    }
}
