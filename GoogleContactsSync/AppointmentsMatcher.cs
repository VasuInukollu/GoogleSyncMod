using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod
{
    internal static class AppointmentsMatcher
    {
        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// Less than 60 seconds doesn't make sense, as the lastSync is saved without seconds and if it is compared
        /// with the LastUpdate dates of Google and Outlook, in the worst case you compare e.g. 15:59 with 16:00 and 
        /// after truncating to minutes you compare 15:00 wiht 16:00
        /// </summary>
        public static int TimeTolerance = 60;

        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        /// <summary>
        /// Matches outlook and Google appointment by a) id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <returns>Returns a list of match pairs (outlook appointment + Google appointment) for all appointment. Those that weren't matche will have it's peer set to null</returns>
        public static List<AppointmentMatch> MatchAppointments(AppointmentsSynchronizer sync)
        {
            Logger.Log("Matching Outlook and Google appointments...", EventType.Information);
            var result = new List<AppointmentMatch>();

            var googleAppointmentExceptions = new List<Event>();

            //for each outlook appointment try to get Google appointment id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without Google appointment. 
            //foreach (Outlook._AppointmentItem olc in outlookAppointments)
            var OutlookAppointmentsWithoutSyncId = new Collection<Outlook.AppointmentItem>();
            #region Match first all outlookAppointments by sync id
            for (int i = 1; i <= sync.OutlookAppointments.Count; i++)
            {
                Outlook.AppointmentItem ola = null;

                try
                {
                    ola = sync.OutlookAppointments[i] as Outlook.AppointmentItem;

                    if (ola == null || string.IsNullOrEmpty(ola.Subject) && ola.Start == AppointmentSync.outlookDateMin)
                    {
                        Logger.Log("Empty Outlook appointment found. Skipping", EventType.Warning);
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }
                    else if (ola.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled || ola.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                    {
                        Logger.Log("Skipping Outlook appointment found because it is cancelled: " + ola.Subject + " - " + ola.Start, EventType.Debug);
                        //sync.SkippedCount++;
                        //sync.SkippedCountNotMatches++;
                        continue;
                    }
                    else if (AppointmentsSynchronizer.TimeMin != null &&
                             (ola.IsRecurring && ola.GetRecurrencePattern().PatternEndDate < AppointmentsSynchronizer.TimeMin ||
                             !ola.IsRecurring && ola.End < AppointmentsSynchronizer.TimeMin) ||
                        AppointmentsSynchronizer.TimeMax != null &&
                             (ola.IsRecurring && ola.GetRecurrencePattern().PatternStartDate > AppointmentsSynchronizer.TimeMax ||
                             !ola.IsRecurring && ola.Start > AppointmentsSynchronizer.TimeMax))
                    {
                        Logger.Log("Skipping Outlook appointment because it is out of months range to sync:" + ola.Subject + " - " + ola.Start, EventType.Debug);
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some appointments throw exceptions
                    if (ola != null && !string.IsNullOrEmpty(ola.Subject))
                        Logger.Log("Accessing Outlook appointment: " + ola.Subject + " threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    else
                        Logger.Log("Accessing Outlook appointment threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                NotificationReceived?.Invoke(string.Format("Matching appointment {0} of {1} by id: {2} ...", i, sync.OutlookAppointments.Count, ola.Subject));

                // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                //OutlookAppointmentInfo olci = new OutlookAppointmentInfo(ola, sync);

                //try to match this appointment to one of Google appointments               

                var gid = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, ola);

                if (gid != null)
                {
                    var e = sync.GetGoogleAppointmentById(gid);
                    var match = new AppointmentMatch(ola, null);

                    if (e != null && !e.Status.Equals("cancelled"))
                    {
                        //we found a match by google id, that is not deleted or cancelled yet
                        match.AddGoogleAppointment(e);
                        result.Add(match);
                        sync.GoogleAppointments.Remove(e);
                    }
                    else
                    {
                        OutlookAppointmentsWithoutSyncId.Add(ola);
                    }
                }
                else
                    OutlookAppointmentsWithoutSyncId.Add(ola);
            }
            #endregion
            #region Match the remaining appointments by properties

            for (int i = 0; i < OutlookAppointmentsWithoutSyncId.Count; i++)
            {
                var ola = OutlookAppointmentsWithoutSyncId[i];

                NotificationReceived?.Invoke(string.Format("Matching appointment {0} of {1} by unique properties: {2} ...", i + 1, OutlookAppointmentsWithoutSyncId.Count, ola.Subject));

                //no match found by id => match by subject/title
                //create a default match pair with just outlook appointment.
                var match = new AppointmentMatch(ola, null);

                //foreach Google appointment try to match and create a match pair if found some match(es)
                for (int j = sync.GoogleAppointments.Count - 1; j >= 0; j--)
                {
                    var e = sync.GoogleAppointments[j];
                    // only match if there is a appointment targetBody, else
                    // a matching Google appointment will be created at each sync                
                    if (!e.Status.Equals("cancelled") && ola.Subject == e.Summary && e.Start.DateTime != null && ola.Start == e.Start.DateTime)
                    {
                        match.AddGoogleAppointment(e);
                        sync.GoogleAppointments.Remove(e);
                    }
                }

                if (match.GoogleAppointment == null)
                    Logger.Log(string.Format("No match found for outlook appointment ({0}) => {1}", match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, (AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment) != null ? "Delete from Outlook" : "Add to Google")), EventType.Information);

                result.Add(match);
            }
            #endregion


            //for each Google appointment that's left (they will be nonmatched) create a new match pair without outlook appointment. 
            for (int i = 0; i < sync.GoogleAppointments.Count; i++)
            {
                var googleAppointment = sync.GoogleAppointments[i];

                NotificationReceived?.Invoke(string.Format("Adding new Google appointment {0} of {1} by unique properties: {2} ...", i + 1, sync.GoogleAppointments.Count, googleAppointment.Summary));

                if (googleAppointment.RecurringEventId != null)
                {
                    sync.SkippedCountNotMatches++;
                    googleAppointmentExceptions.Add(googleAppointment);
                }
                else if (googleAppointment.Status.Equals("cancelled"))
                {
                    Logger.Log("Skipping Google appointment found because it is cancelled: " + googleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(googleAppointment), EventType.Debug);
                    //sync.SkippedCount++;
                    //sync.SkippedCountNotMatches++;
                }
                else if (string.IsNullOrEmpty(googleAppointment.Summary) && (googleAppointment.Start == null || googleAppointment.Start.DateTime == null && googleAppointment.Start.Date == null))
                {
                    // no title or time
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    Logger.Log("Skipped GoogleAppointment because no unique property found (Subject or StartDate):" + googleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(googleAppointment), EventType.Warning);
                }
                else
                {
                    Logger.Log(string.Format("No match found for Google appointment ({0}) => {1}", googleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(googleAppointment), (!string.IsNullOrEmpty(AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, googleAppointment)) ? "Delete from Google" : "Add to Outlook")), EventType.Information);
                    var match = new AppointmentMatch(null, googleAppointment);
                    result.Add(match);
                }
            }

            //for each Google appointment exception, assign to proper match
            for (int i = 0; i < googleAppointmentExceptions.Count; i++)
            {
                var e = googleAppointmentExceptions[i];
                NotificationReceived?.Invoke(string.Format("Adding Google appointment exception {0} of {1} : {2} ...", i + 1, googleAppointmentExceptions.Count, e.Summary + " - " + AppointmentsSynchronizer.GetTime(e)));

                //Search for original recurrent event in matches
                bool found = false;
                foreach (var m in result)
                {
                    if (m.GoogleAppointment != null && e.RecurringEventId.Equals(m.GoogleAppointment.Id))
                    {
                        m.GoogleAppointmentExceptions.Add(e);
                        found = true;
                        break;
                    }
                }

                if (!found)
                    Logger.Log(string.Format("No match found for Google appointment exception: {0}", e.Summary + " - " + AppointmentsSynchronizer.GetTime(e)), EventType.Debug);
            }

            return result;
        }

        public static void SyncAppointments(AppointmentsSynchronizer sync)
        {
            for (int i = 0; i < sync.Appointments.Count; i++)
            {
                AppointmentMatch match = sync.Appointments[i];
                if (NotificationReceived != null)
                {
                    string name = string.Empty;
                    if (match.OutlookAppointment != null)
                        name = match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start;
                    else if (match.GoogleAppointment != null)
                        name = match.GoogleAppointment.Summary + " - " + AppointmentsSynchronizer.GetTime(match.GoogleAppointment);
                    NotificationReceived(string.Format("Syncing appointment {0} of {1}: {2} ...", i + 1, sync.Appointments.Count, name));
                }

                SyncAppointment(match, sync);
            }
        }

        private static void SyncAppointmentNoGoogle(AppointmentMatch match, AppointmentsSynchronizer sync)
        {
            string googleAppointmentId = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
            if (!string.IsNullOrEmpty(googleAppointmentId))
            {
                //Redundant check if exist, but in case an error occurred in MatchAppointments or not all appointments have been loaded (e.g. because months before/after constraint)
                Event matchingGoogleAppointment = null;
                if (sync.AllGoogleAppointments != null)
                    matchingGoogleAppointment = sync.GetGoogleAppointmentById(googleAppointmentId);
                else
                    matchingGoogleAppointment = sync.LoadGoogleAppointments(googleAppointmentId, null, null, null, null);
                if (matchingGoogleAppointment == null)
                {
                    if (sync.SyncOption == SyncOption.OutlookToGoogleOnly || !sync.SyncDelete)
                        return;
                    else if (!sync.PromptDelete && match.OutlookAppointment.Recipients.Count == 0)
                        sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
                    else if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                             sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                    {
                        using (var r = new ConflictResolver())
                        {
                            sync.DeleteOutlookResolution = r.ResolveDelete(match.OutlookAppointment);
                        }
                    }
                    switch (sync.DeleteOutlookResolution)
                    {
                        case DeleteResolution.KeepOutlook:
                        case DeleteResolution.KeepOutlookAlways:
                            AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                            break;
                        case DeleteResolution.DeleteOutlook:
                        case DeleteResolution.DeleteOutlookAlways:

                            if (match.OutlookAppointment.Recipients.Count > 1)
                            {
                                //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment                                
                                Logger.Log("Outlook Appointment not deleted, because multiple participants found,  invitation maybe NOT sent by Google: " + match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, EventType.Information);
                                AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                                break;
                            }
                            else
                                //Avoid recreating a GoogleAppointment already existing
                                //==> Delete this OutlookAppointment instead if previous match existed but no match exists anymore
                                return;
                        default:
                            throw new ApplicationException("Cancelled");
                    }
                }
                else
                {
                    sync.SkippedCount++;
                    match.GoogleAppointment = matchingGoogleAppointment;
                    Logger.Log("Outlook Appointment not deleted, because still existing on Google side, maybe because months restriction: " + match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, EventType.Information);
                    return;
                }
            }

            if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
            {
                sync.SkippedCount++;
                Logger.Log(string.Format("Outlook appointment not added to Google, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.OutlookAppointment.Subject), EventType.Information);
                return;
            }

            //create a Google appointment from Outlook appointment
            match.GoogleAppointment = Factory.NewEvent();

            sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);
        }

        private static void SyncAppointmentNoOutlook(AppointmentMatch match, AppointmentsSynchronizer sync)
        {
            string outlookAppointmenttId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, match.GoogleAppointment);
            if (!string.IsNullOrEmpty(outlookAppointmenttId))
            {
                if (sync.SyncOption == SyncOption.GoogleToOutlookOnly || !sync.SyncDelete)
                    return;
                else if (!sync.PromptDelete)
                    sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                         sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                {
                    using (var r = new ConflictResolver())
                    {
                        sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleAppointment);
                    }
                }
                switch (sync.DeleteGoogleResolution)
                {
                    case DeleteResolution.KeepGoogle:
                    case DeleteResolution.KeepGoogleAlways:
                        AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(sync.SyncProfile, match.GoogleAppointment);
                        break;
                    case DeleteResolution.DeleteGoogle:
                    case DeleteResolution.DeleteGoogleAlways:
                        //Avoid recreating a OutlookAppointment already existing
                        //==> Delete this e instead if previous match existed but no match exists anymore 
                        return;
                    default:
                        throw new ApplicationException("Cancelled");
                }
            }


            if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
            {
                sync.SkippedCount++;
                Logger.Log(string.Format("Google appointment not added to Outlook, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.GoogleAppointment.Summary), EventType.Information);
                return;
            }

            //create a Outlook appointment from Google appointment
            match.OutlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);

            sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
        }

        private static void SyncAppointmentBothExists(AppointmentMatch m, AppointmentsSynchronizer sync)
        {
            //ToDo: Check how to overcome appointment recurrences, which need more than 60 seconds to update and therefore get updated again and again because of time tolerance 60 seconds violated again and again

            //merge appointment details                

            //determine if this appointment pair were synchronized
            //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookAppointment, sync.OutlookPropertyNameUpdated);
            DateTime? lastSynced = AppointmentPropertiesUtils.GetOutlookLastSync(sync, m.OutlookAppointment);
            if (lastSynced.HasValue)
            {
                //appointment pair was syncronysed before.

                //determine if Google appointment was updated since last sync

                //lastSynced is stored without seconds. take that into account.
                DateTime lastUpdatedOutlook = m.OutlookAppointment.LastModificationTime.AddSeconds(-m.OutlookAppointment.LastModificationTime.Second);
                DateTime lastUpdatedGoogle = m.GoogleAppointment.Updated.Value.AddSeconds(-m.GoogleAppointment.Updated.Value.Second);
                //consider GoogleAppointmentExceptions, because if they are updated, the master appointment doesn't have a new Saved TimeStamp
                foreach (var e in m.GoogleAppointmentExceptions)
                {
                    if (e.Updated != null)//happens for cancelled events
                    {
                        DateTime lastUpdatedGoogleException = e.Updated.Value.AddSeconds(-e.Updated.Value.Second);
                        if (lastUpdatedGoogleException > lastUpdatedGoogle)
                            lastUpdatedGoogle = lastUpdatedGoogleException;
                    }
                    else if (m.OutlookAppointment.IsRecurring && m.OutlookAppointment.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                    {
                        Outlook.AppointmentItem outlookRecurrenceException = null;
                        try
                        {
                            var slaveRecurrence = m.OutlookAppointment.GetRecurrencePattern();
                            if (e.OriginalStartTime != null && !string.IsNullOrEmpty(e.OriginalStartTime.Date))
                                outlookRecurrenceException = slaveRecurrence.GetOccurrence(DateTime.Parse(e.OriginalStartTime.Date));
                            else if (e.OriginalStartTime != null && e.OriginalStartTime.DateTime != null)
                                outlookRecurrenceException = slaveRecurrence.GetOccurrence(e.OriginalStartTime.DateTime.Value);
                        }
                        catch (Exception ignored)
                        {
                            Logger.Log("Google Appointment with OriginalEvent found, but Outlook occurrence not found: " + e.Summary + " - " + e.OriginalStartTime.DateTime + ": " + ignored, EventType.Debug);
                        }

                        if (outlookRecurrenceException != null && outlookRecurrenceException.MeetingStatus != Outlook.OlMeetingStatus.olMeetingCanceled)
                        {
                            lastUpdatedGoogle = DateTime.Now;
                            break; //no need to search further, already newest date set
                        }
                    }
                }

                //check if both outlok and Google appointments where updated sync last sync
                if ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance
                    && (int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance)
                {
                    //both appointments were updated.
                    //options: 1) ignore 2) loose one based on SyncOption
                    //throw new Exception("Both appointments were updated!");

                    switch (sync.SyncOption)
                    {
                        case SyncOption.MergeOutlookWins:
                        case SyncOption.OutlookToGoogleOnly:
                            //overwrite Google appointment
                            Logger.Log("Outlook and Google appointment have been updated, Outlook appointment is overwriting Google because of SyncOption " + sync.SyncOption + ": " + m.OutlookAppointment.Subject + ".", EventType.Information);
                            sync.UpdateAppointment(m.OutlookAppointment, ref m.GoogleAppointment);
                            break;
                        case SyncOption.MergeGoogleWins:
                        case SyncOption.GoogleToOutlookOnly:
                            //overwrite outlook appointment
                            Logger.Log("Outlook and Google appointment have been updated, Google appointment is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + m.GoogleAppointment.Summary + ".", EventType.Information);
                            sync.UpdateAppointment(ref m.GoogleAppointment, m.OutlookAppointment, m.GoogleAppointmentExceptions);
                            break;
                        case SyncOption.MergePrompt:
                            //promp for sync option
                            if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.SkipAlways)
                            {
                                using (var r = new ConflictResolver())
                                {
                                    sync.ConflictResolution = r.Resolve(m.OutlookAppointment, m.GoogleAppointment, sync, false);
                                }
                            }
                            switch (sync.ConflictResolution)
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways:
                                    Logger.Log(string.Format("User skipped appointment ({0}).", m.ToString()), EventType.Information);
                                    sync.SkippedCount++;
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways:
                                    sync.UpdateAppointment(m.OutlookAppointment, ref m.GoogleAppointment);
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways:
                                    sync.UpdateAppointment(ref m.GoogleAppointment, m.OutlookAppointment, m.GoogleAppointmentExceptions);
                                    break;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                            break;
                    }
                    return;
                }


                //check if Outlook appointment was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.GoogleToOutlookOnly &&
                    ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                     (int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                     sync.SyncOption == SyncOption.OutlookToGoogleOnly
                    )
                   )
                {
                    //Outlook appointment was changed or changed Google appointment will be overwritten

                    if ((int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                        sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                        Logger.Log("Google appointment has been updated since last sync, but Outlook appointment is overwriting Google because of SyncOption " + sync.SyncOption + ": " + m.OutlookAppointment.Subject + ".", EventType.Information);

                    sync.UpdateAppointment(m.OutlookAppointment, ref m.GoogleAppointment);

                    //at the moment use Outlook as "master" source of appointments - in the event of a conflict Google appointment will be overwritten.
                    //TODO: control conflict resolution by SyncOption
                    return;
                }

                //check if Google appointment was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.OutlookToGoogleOnly &&
                    ((int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                     (int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                     sync.SyncOption == SyncOption.GoogleToOutlookOnly
                    )
                   )
                {
                    //google appointment was changed or changed Outlook appointment will be overwritten

                    if ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                        sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                        Logger.Log("Outlook appointment has been updated since last sync, but Google appointment is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + m.OutlookAppointment.Subject + ".", EventType.Information);

                    sync.UpdateAppointment(ref m.GoogleAppointment, m.OutlookAppointment, m.GoogleAppointmentExceptions);
                }
            }
            else
            {
                //appointments were never synced.
                //merge appointments.
                switch (sync.SyncOption)
                {
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite Google appointment
                        sync.UpdateAppointment(m.OutlookAppointment, ref m.GoogleAppointment);
                        break;
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite outlook appointment
                        sync.UpdateAppointment(ref m.GoogleAppointment, m.OutlookAppointment, m.GoogleAppointmentExceptions);
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                sync.ConflictResolution = r.Resolve(m.OutlookAppointment, m.GoogleAppointment, sync, true);
                            }
                        }
                        switch (sync.ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                sync.Appointments.Add(new AppointmentMatch(m.OutlookAppointment, null));
                                sync.Appointments.Add(new AppointmentMatch(null, m.GoogleAppointment));
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways:
                                sync.UpdateAppointment(m.OutlookAppointment, ref m.GoogleAppointment);
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways:
                                sync.UpdateAppointment(ref m.GoogleAppointment, m.OutlookAppointment, m.GoogleAppointmentExceptions);
                                break;
                            default:
                                throw new ApplicationException("Canceled");
                        }
                        break;
                }
            }
        }

        public static void SyncAppointment(AppointmentMatch m, AppointmentsSynchronizer sync)
        {
            if (m.GoogleAppointment == null && m.OutlookAppointment != null)
            {
                //no Google appointment                               
                SyncAppointmentNoGoogle(m, sync);
            }
            else if (m.OutlookAppointment == null && m.GoogleAppointment != null)
            {
                //no Outlook appointment                               
                SyncAppointmentNoOutlook(m, sync);
            }
            else if (m.OutlookAppointment != null && m.GoogleAppointment != null)
            {
                SyncAppointmentBothExists(m, sync);
            }
            else
                throw new ArgumentNullException("AppointmenttMatch has all peers null.");
        }
    }



    internal class AppointmentMatch
    {
        //ToDo: OutlookappointmentInfo
        public Outlook.AppointmentItem OutlookAppointment;
        public Event GoogleAppointment;
        public readonly List<Event> AllGoogleAppointmentMatches = new List<Event>(1);
        public Event LastGoogleAppointment;
        public List<Event> GoogleAppointmentExceptions = new List<Event>();

        public AppointmentMatch(Outlook.AppointmentItem outlookAppointment, Event googleAppointment)
        {
            OutlookAppointment = outlookAppointment;
            GoogleAppointment = googleAppointment;
        }

        public void AddGoogleAppointment(Event googleAppointment)
        {
            if (googleAppointment == null)
                return;
            //throw new ArgumentNullException("e must not be null.");

            if (GoogleAppointment == null)
                GoogleAppointment = googleAppointment;

            //this to avoid searching the entire collection. 
            //if last appointment it what we are trying to add the we have already added it earlier
            if (LastGoogleAppointment == googleAppointment)
                return;

            if (!AllGoogleAppointmentMatches.Contains(googleAppointment))
                AllGoogleAppointmentMatches.Add(googleAppointment);

            LastGoogleAppointment = googleAppointment;
        }

    }



}
