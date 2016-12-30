using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.GData.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    class AppointmentsSynchronizer : IDisposable
    {
        public Outlook.Items OutlookAppointments { get; set; }
        public Collection<Event> GoogleAppointments { get; set; }
        public Collection<Event> AllGoogleAppointments { get; set; }

        public static bool SyncAppointmentsForceRTF { get; set; }

        public EventsResource EventRequest { get; set; }

        public static string SyncAppointmentsFolder { get; set; }
        public static string SyncAppointmentsGoogleFolder { get; set; }
        public static string SyncAppointmentsGoogleTimeZone { get; set; }

        public static DateTime? TimeMin { get; set; }
        public static DateTime? TimeMax { get; set; }

        public static string Timezone { get; set; }
        public static bool MappingBetweenTimeZonesRequired { get; set; }

        public CalendarService CalendarRequest;

        public IList<CalendarListEntry> calendarList { get; set; }
        public ConflictResolution ConflictResolution { get; set; }
        public DeleteResolution DeleteGoogleResolution { get; set; }
        public DeleteResolution DeleteOutlookResolution { get; set; }

        public string OutlookPropertyNameId
        {
            get { return OutlookPropertyPrefix + "id"; }
        }
        public string OutlookPropertyPrefix { get; set; }
        public string OutlookPropertyNameSynced
        {
            get { return OutlookPropertyPrefix + "up"; }
        }

        public event ErrorNotificationHandler ErrorEncountered;
        public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
        public string SyncProfile { get; set; }
        public bool SyncDelete { get; set; }
        public bool PromptDelete { get; set; }
        public int SkippedCountNotMatches { get; set; }
        public int DeletedCount { get; set; }
        public int SyncedCount { get; set; }
        public int ErrorCount { get; set; }
        public int SkippedCount { get; set; }


        public List<AppointmentMatch> Appointments { get; set; }

        private SyncOption _syncOption = SyncOption.MergeOutlookWins;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; }
        }

        public void ReadGoogleAppointmentConfig(Google.Apis.Services.BaseClientService.Initializer initializer)
        {
            CalendarRequest = GoogleServices.CreateCalendarService(initializer);



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
                        Logger.Log("Multiple participants found, invitation maybe NOT sent by Google. Outlook appointment is overwriting Google because of SyncOption " + SyncOption + ": " + master.Summary + " - " + AppointmentsSynchronizer.GetTime(master) + ". ", EventType.Information);
                        UpdateAppointment(slave, ref master);
                        break;
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite outlook appointment
                        Logger.Log("Multiple participants found, invitation maybe NOT sent by Google, but Google appointment is overwriting Outlook because of SyncOption " + SyncOption + ": " + master.Summary + " - " + AppointmentsSynchronizer.GetTime(master) + ".", EventType.Information);
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
                                ConflictResolution = r.Resolve("Cannot update appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + AppointmentsSynchronizer.GetTime(master) + "\". Do you want to update it back from Outlook to Google?", slave, master, this);
                            }
                        }
                        switch (ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Skip
                                SkippedCount++;
                                Logger.Log("Skipped Updating appointment from Google to Outlook because multiple participants found, invitation maybe NOT sent by Google: \"" + master.Summary + " - " + AppointmentsSynchronizer.GetTime(master) + "\".", EventType.Information);
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
        /// Resets associations of Outlook appointments with Google appointments via user props
        /// and vice versa
        /// </summary>
        public void ResetOutlookAppointmentMatches(bool deleteOutlookAppointments)
        {
            Debug.Assert(OutlookAppointments != null, "Outlook Appointments object is null - this should not happen. Please inform Developers.");

            //try
            //{

            lock (Synchronizer._syncRoot)
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
        /// Reset the match link between Outlook and Google appointment
        /// </summary>
        public void ResetMatch(Outlook.AppointmentItem ola)
        {
            if (ola != null)
            {
                AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, ola);
                ola.Save();
            }
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

        public static Outlook.AppointmentItem CreateOutlookAppointmentItem(string syncAppointmentsFolder)
        {
            //OutlookAppointment = OutlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.AppointmentItem outlookAppointment = null;
            Outlook.MAPIFolder appointmentsFolder = null;
            Outlook.Items items = null;

            try
            {
                appointmentsFolder = Synchronizer.OutlookNameSpace.GetFolderFromID(syncAppointmentsFolder);
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

                if (TimeMin != null)
                    query.TimeMin = TimeMin;
                if (TimeMax != null)
                    query.TimeMax = TimeMax;

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

                if (TimeMin != null)
                    query.TimeMin = TimeMin;
                if (TimeMax != null)
                    query.TimeMax = TimeMax;

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

        /// <summary>
        /// Load the appointments from Google and Outlook and match them
        /// </summary>
        public void MatchAppointments()
        {
            LoadAppointments();
            Appointments = AppointmentsMatcher.MatchAppointments(this);
            Logger.Log("Appointment Matches Found: " + Appointments.Count, EventType.Debug);
        }

        public void LoadAppointments()
        {
            LoadOutlookAppointments();
            LoadGoogleAppointments();
            RemoveOutlookDuplicatedAppointments();
            RemoveGoogleDuplicatedAppointments();
        }

        internal Google.Apis.Calendar.v3.Data.Event LoadGoogleAppointments(string id, DateTime? start, DateTime? end, DateTime? restrictStartTime, DateTime? restrictEndTime)
        {
            string message = "Error Loading Google appointments. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Google.Apis.Calendar.v3.Data.Event ret = null;
            try
            {

                GoogleAppointments = new Collection<Google.Apis.Calendar.v3.Data.Event>();

                var query = EventRequest.List(SyncAppointmentsGoogleFolder);

                string pageToken = null;
                //query.MaxResults = 256; //ToDo: Find a way to retrieve all appointments

                //Only Load events from month range, but only if not a distinct Google Appointment is searched for
                if (start != null)
                    query.TimeMin = TimeMin;
                if (restrictStartTime != null && (query.TimeMin == default(DateTime) || restrictStartTime > query.TimeMin))
                    query.TimeMin = restrictStartTime.Value;
                if (end != null)
                    query.TimeMax = TimeMax;
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
            if (start == null && end == null && restrictStartTime == null && restrictEndTime == null) //restrictStartDate == null)
                AllGoogleAppointments = GoogleAppointments;

            return ret;
        }

        private void LoadGoogleAppointments()
        {
            Logger.Log("Loading Google appointments...", EventType.Information);
            LoadGoogleAppointments(null, TimeMin, TimeMax, null, null);
            Logger.Log("Google Appointments Found: " + GoogleAppointments.Count, EventType.Debug);
        }

        /// <summary>
        /// Remove duplicates from Google: two different Google appointments pointing to the same Outlook appointment.
        /// </summary>
        private void RemoveGoogleDuplicatedAppointments()
        {
            Logger.Log("Removing Google duplicated appointments...", EventType.Information);

            if (GoogleAppointments.Count < 2)
                return;

            var appointments = new Dictionary<string, int>();

            //scan all Google appointments
            for (int i = 0; i < GoogleAppointments.Count; i++)
            {
                var e1 = GoogleAppointments[i];
                if (e1 == null)
                    continue;

                try
                {
                    var oid = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, e1);

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
                                var gid = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, a);

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
                                a = null;
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

        private void LoadOutlookAppointments()
        {
            Logger.Log("Loading Outlook appointments...", EventType.Information);
            OutlookAppointments = Synchronizer.GetOutlookItems(Outlook.OlDefaultFolders.olFolderCalendar, SyncAppointmentsFolder);
            Logger.Log("Outlook Appointments Found: " + OutlookAppointments.Count, EventType.Debug);
        }

        /// <summary>
        /// Remove duplicates from Outlook: two different Outlook appointments pointing to the same Google appointment.
        /// Such situation typically happens when copy/paste'ing synchronized appointment in Outlook
        /// </summary>
        private void RemoveOutlookDuplicatedAppointments()
        {
            Logger.Log("Removing Outlook duplicated appointments...", EventType.Information);

            if (OutlookAppointments.Count < 2)
                return;

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

                    var gid = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, ola1);
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
                                var oid = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, e);
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
                            {
                                Marshal.ReleaseComObject(ola2);
                                ola2 = null;
                            }
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
                    {
                        Marshal.ReleaseComObject(ola1);
                        ola1 = null;
                    }
                }
            }
        }

        public void Dispose()
        {
            ((IDisposable)CalendarRequest).Dispose();
        }
    }
}
