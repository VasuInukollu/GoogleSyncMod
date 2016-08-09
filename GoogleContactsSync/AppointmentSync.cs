using System;
using System.Collections.Generic;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;
using System.Linq;

namespace GoContactSyncMod
{

    internal static class AppointmentSync
    {
        private const string dateFormat = "yyyyMMdd";
        private const string timeFormat = "HHmmss";
        internal static DateTime outlookDateMin = new DateTime(4501, 1, 1);
        internal static DateTime outlookDateMax = new DateTime(4500, 12, 31);

        const string DTSTART = "DTSTART";
        const string DTEND = "DTEND";
        const string RRULE = "RRULE";
        const string FREQ = "FREQ";
        const string DAILY = "DAILY";
        const string WEEKLY = "WEEKLY";
        const string MONTHLY = "MONTHLY";
        const string YEARLY = "YEARLY";

        const string BYMONTH = "BYMONTH";
        const string BYMONTHDAY = "BYMONTHDAY";
        const string BYDAY = "BYDAY";
        const string BYSETPOS= "BYSETPOS";

        const string VALUE = "VALUE";
        const string DATE = "DATE";
        const string DATETIME = "DATE-TIME";
        const string INTERVAL = "INTERVAL";
        const string COUNT = "COUNT";
        const string UNTIL = "UNTIL";
        const string TZID = "TZID";

        const string MO = "MO";
        const string TU = "TU";
        const string WE = "WE";
        const string TH = "TH";
        const string FR = "FR";
        const string SA = "SA";
        const string SU = "SU";

        // This will return the Windows zone that matches the IANA zone, if one exists.
        private static string IanaToWindows(string ianaZoneId)
        {
            var utcZones = new[] { "Etc/UTC", "Etc/UCT", "Etc/GMT" };
            if (utcZones.Contains(ianaZoneId, StringComparer.Ordinal))
                return "UTC";

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;

            // resolve any link, since the CLDR doesn't necessarily use canonical IDs
            var links = tzdbSource.CanonicalIdMap
                .Where(x => x.Value.Equals(ianaZoneId, StringComparison.Ordinal))
                .Select(x => x.Key);

            // resolve canonical zones, and include original zone as well
            var possibleZones = tzdbSource.CanonicalIdMap.ContainsKey(ianaZoneId)
                ? links.Concat(new[] { tzdbSource.CanonicalIdMap[ianaZoneId], ianaZoneId })
                : links;

            // map the windows zone
            var mappings = tzdbSource.WindowsMapping.MapZones;
            var item = mappings.FirstOrDefault(x => x.TzdbIds.Any(possibleZones.Contains));
            if (item == null) return null;
            return item.WindowsId;
        }

        // This will return the "primary" IANA zone that matches the given windows zone.
        // If the primary zone is a link, it then resolves it to the canonical ID.
        public static string WindowsToIana(string windowsZoneId)
        {
            if (windowsZoneId.Equals("UTC", StringComparison.Ordinal))
                return "Etc/UTC";

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;

            TimeZoneInfo tzi = null;
            try
            {
                tzi = TimeZoneInfo.FindSystemTimeZoneById(windowsZoneId);
            }
            catch (Exception)
            {
                return null;
            }
            if (tzi == null) return null;
            var tzid = tzdbSource.MapTimeZoneId(tzi);
            if (tzid == null) return null;
            return tzdbSource.CanonicalIdMap[tzid];
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void UpdateAppointment(Outlook.AppointmentItem master, Event slave)
        {            
            slave.Summary = master.Subject;            

            ////foreach (Outlook.Attachment attachment in master.Attachments)
            ////    slave.Attachments.Add(master.Attachments);
            
            slave.Description = master.Body;            

            switch (master.BusyStatus)
            {
                case Outlook.OlBusyStatus.olBusy:
                    slave.Transparency = "opaque";
                    slave.Status = "confirmed";
                    break;
                case Outlook.OlBusyStatus.olTentative:
                    slave.Transparency = "transparent";
                    slave.Status = "tentative";
                    break;
                case Outlook.OlBusyStatus.olOutOfOffice:
                    slave.Transparency = "opaque";
                    slave.Status = "tentative";
                    break;
                //ToDo: case Outlook.OlBusyStatus.olWorkingElsewhere:
                case Outlook.OlBusyStatus.olFree:                
                default:
                    slave.Status = "confirmed";
                    slave.Transparency = "transparent";
                    break;
            }

            //ToDo:slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            slave.Location = master.Location;


            //if (!master.IsRecurring || master.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster)
            //{                   
                if (master.AllDayEvent == true) 
                {
                    slave.Start.Date = master.Start.ToString("yyyy-MM-dd");
                    slave.End.Date = master.End.ToString("yyyy-MM-dd");
                    slave.Start.DateTime = null;
                    slave.End.DateTime = null;
                } 
                else 
                {
                    slave.Start.Date = null;
                    slave.End.Date = null;
                    slave.Start.DateTime = master.Start;
                    slave.End.DateTime = master.End;
                    if (master.StartTimeZone != null && !string.IsNullOrEmpty(master.StartTimeZone.ID))
                        slave.Start.TimeZone = WindowsToIana(master.StartTimeZone.ID);
                   if (master.EndTimeZone != null && !string.IsNullOrEmpty(master.EndTimeZone.ID))
                        slave.End.TimeZone = WindowsToIana(master.EndTimeZone.ID);
                }

                if (master.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                {   //As Timezone is only mandatory for recurrence events
                    //If user has selected a timezone in settings form, use this as timezone, otherwise assign the Outlook timezoen (not working always and sometimes not compatible
                    //A list of TZ can be found here: http://en.wikipedia.org/wiki/List_of_tz_database_time_zones
                    if (!string.IsNullOrEmpty(Synchronizer.Timezone))
                    {
                        slave.Start.TimeZone = Synchronizer.Timezone;
                        slave.End.TimeZone = Synchronizer.Timezone;
                    }
                    else
                    {
                        if (master.StartTimeZone != null && !string.IsNullOrEmpty(master.StartTimeZone.Name))
                            slave.Start.TimeZone = master.StartTimeZone.Name;                       
                        if (master.EndTimeZone != null && !string.IsNullOrEmpty(master.EndTimeZone.Name))
                            slave.End.TimeZone = master.EndTimeZone.Name;
                    }
                }            
            //}


            #region participants
            //ToDo: Commented out for now, not sync participants, because otherwise Google raises quota exceptions
            //slave.Participants.Clear();
            //int i = 0;
            //foreach (Outlook.Recipient recipient in master.Recipients)
            //{
             
            //    var participant = new EventAttendee();

            //    participant.Email = AppointmentPropertiesUtils.GetOutlookEmailAddress(master.Subject + " - " + master.Start, recipient);

            //    participant.Rel = (i == 0 ? EventAttendee.RelType.EVENT_ORGANIZER : EventAttendee.RelType.EVENT_ATTENDEE);
            //    slave.Participants.Add(participant);
            //    i++;
            //}
            //End Todo commented out
            #endregion

            //slave.RequiredAttendees = master.RequiredAttendees;
            //slave.OptionalAttendees = master.OptionalAttendees;

            //ToDo: Doesn'T work for newly created appointments, because Event.Reminder is throwing NullPointerException and Reminders cannot be initialized, therefore moved to after saving
            //if (slave.Reminders != null && slave.Reminders.Overrides != null)
            //{
            //    slave.Reminders.Overrides.Clear();
            //    if (master.ReminderSet)
            //    {
            //        var reminder = new EventReminder();
            //        reminder.Minutes = master.ReminderMinutesBeforeStart;
            //        if (reminder.Minutes > 40300)
            //        {
            //            //ToDo: Check real limit, currently 40300
            //            Logger.Log("Reminder Minutes to big (" + reminder.Minutes + "), set to maximum of 40300 minutes for appointment: " + master.Subject + " - " + master.Start, EventType.Warning);
            //            reminder.Minutes = 40300;                        
            //        }
            //        reminder.Method = "popup";
            //        slave.Reminders.Overrides.Add(reminder);
            //    }
            //}
            UpdateAppointmentReminders(master, slave);

            //slave.Resources = master.Resources;

            UpdateRecurrence(master, slave);

            if (slave.Recurrence == null)
            {
                switch (master.Sensitivity)
                {
                    case Outlook.OlSensitivity.olConfidential: //ToDo, currently not supported by Google Web App GUI and Outlook 2010: slave.EventVisibility = Google.GData.Calendar.Event.Visibility.CONFIDENTIAL; break;#
                    case Outlook.OlSensitivity.olPersonal: //ToDo, currently not supported by Google Web App GUI and Outlook 2010: slave.EventVisibility = Google.GData.Calendar.Event.Visibility.CONFIDENTIAL; break;
                    case Outlook.OlSensitivity.olPrivate: slave.Visibility = "private"; break;
                    default: slave.Visibility = "default"; break;
                }
            }
        }

        public static void UpdateAppointmentReminders(Outlook.AppointmentItem master, Event slave)
       {
            if (master.ReminderSet)
            {
                if (slave.Reminders == null)
                {
                    slave.Reminders = new Event.RemindersData();
                    slave.Reminders.Overrides = new List<EventReminder>();
                }

                slave.Reminders.UseDefault = false;
                if (slave.Reminders.Overrides != null)
                {
                    slave.Reminders.Overrides.Clear();
                }
                else
                {
                    slave.Reminders.Overrides = new List<EventReminder>();
                }
                var reminder = new EventReminder();
                reminder.Minutes = master.ReminderMinutesBeforeStart;
                if (reminder.Minutes > 40300)
                {
                    //ToDo: Check real limit, currently 40300
                    Logger.Log("Reminder Minutes to big ("  + reminder.Minutes + "), set to maximum of 40300 minutes for appointment: "  + master.Subject +   " - "  + master.Start, EventType.Warning);
                    reminder.Minutes = 40300;
                }
                reminder.Method = "popup";
                slave.Reminders.Overrides.Add(reminder);
            }
            else if (slave.Reminders != null)
            {
                if (slave.Reminders.Overrides != null)
                {
                    slave.Reminders.Overrides.Clear();
                }
                slave.Reminders.UseDefault = false;
            }           
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void UpdateAppointment(Event master, Outlook.AppointmentItem slave)
        {
            slave.Subject = master.Summary;            

            //foreach (Outlook.Attachment attachment in master.Attachments)
            //    slave.Attachments.Add(master.Attachments);

            try
            {
                string nonRTF;
                if (slave.Body == null)
                    nonRTF = string.Empty;
                else
                    nonRTF = Utilities.ConvertToText(slave.RTFBody as byte[]);

                if (string.IsNullOrEmpty(nonRTF) || nonRTF.Equals(slave.Body) && !nonRTF.Equals(master.Description))
                {  //only update, if RTF text is same as plain text and is different between master and slave
                    slave.Body = master.Description;
                }
                else if (!nonRTF.Equals(master.Description))
                {
                    if (!Synchronizer.SyncAppointmentsForceRTF)
                    {
                        slave.Body = master.Description;
                    }
                    else
                    {
                        Logger.Log("Outlook appointment notes body not updated, because it is RTF, otherwise it will overwrite it by plain text: " + slave.Subject + " - " + slave.Start, EventType.Warning);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Log("Error when converting RTF to plain text, updating Google Appointment '"+ slave.Subject + " - " + slave.Start +"' notes to Outlook without RTF check: " + e.Message , EventType.Debug);
                slave.Body = master.Description;
            }
                                                                   

            //slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            slave.Location = master.Location;                    

            //if (master.Times.Count != 1 && master.Recurrence == null)
            //    Logger.Log("Google Appointment with multiple or no times found: " + master.Summary + " - " + Syncronizer.GetTime(master), EventType.Warning);

            //if (master.RecurrenceException != null)
            //    Logger.Log("Google Appointment with RecurrenceException found: " + master.Summary + " - " + Syncronizer.GetTime(master), EventType.Warning);            

            //if (master.Times.Count == 1 || master.Times.Count > 0 && master.Recurrence == null)
            //if (master.Times.Count > 0)
            {//Also sync times for recurrent events, but log warning, if it is not working
                try
                {
                    if (master.Start != null && slave.AllDayEvent == string.IsNullOrEmpty(master.Start.Date))
                        slave.AllDayEvent = !string.IsNullOrEmpty(master.Start.Date);
                    if (master.Start != null && !string.IsNullOrEmpty(master.Start.Date))
                        slave.Start = DateTime.Parse(master.Start.Date);
                    else if (master.Start != null && master.Start.DateTime != null && slave.Start != master.Start.DateTime)
                        slave.Start = master.Start.DateTime.Value;
                    if (master.End != null && !string.IsNullOrEmpty(master.End.Date))
                        slave.End = DateTime.Parse(master.End.Date);
                    else if (master.End != null && master.End.DateTime != null && slave.End != master.End.DateTime)
                        slave.End = master.End.DateTime.Value;
                }
                catch (Exception ex)
                {
                    Logger.Log("Error updating event's AllDay/Start/End: " + master.Summary + " - " + Synchronizer.GetTime(master) + ": " + ex.Message, slave.IsRecurring?EventType.Debug:EventType.Warning);
                }
            }
            
            //slave.StartInStartTimeZone = master.StartInStartTimeZone;
            //slave.StartTimeZone = master.StartTimeZone;
            //slave.StartUTC = master.StartUTC;

            //if (!IsOrganizer(GetOrganizer(master)) || !IsOrganizer(GetOrganizer(slave), slave))
            //    slave.MeetingStatus = Outlook.OlMeetingStatus.olMeetingReceived;

	        try
	        {
		        if (master.Status.Equals("confirmed") && (master.Transparency == null || master.Transparency.Equals("opaque")))
			        slave.BusyStatus = Outlook.OlBusyStatus.olBusy;
		        else if ((master.Status.Equals("confirmed") && master.Transparency.Equals("transparent")) || master.Status.Equals("cancelled"))
			        slave.BusyStatus = Outlook.OlBusyStatus.olFree;
		        else if (master.Status.Equals("tentative") && (master.Transparency == null || master.Transparency.Equals("opaque"))) 
			        slave.BusyStatus = Outlook.OlBusyStatus.olOutOfOffice;
                else if (master.Status.Equals("tentative"))
                    slave.BusyStatus = Outlook.OlBusyStatus.olTentative;
                else
			        slave.BusyStatus = Outlook.OlBusyStatus.olWorkingElsewhere;
	        }
	        catch (Exception e)
	        {
		        MessageBox.Show(e.StackTrace);
		        throw;
	        }

	        #region Recipients
            //ToDo: Commented out for now, not sync participants, because otherwise Google raises quota exceptions
            //for (int i = slave.Recipients.Count; i > 0; i--)
            //    slave.Recipients.Remove(i);


            ////Add Organizer
            //foreach (EventAttendee participant in master.Participants)
            //{
            //    if (participant.Rel == EventAttendee.RelType.EVENT_ORGANIZER && participant.Email != Syncronizer.UserName)
            //    {
            //        //ToDo: Doesn't Work, because Organizer cannot be set on Outlook side (it is ignored)
            //        //slave.GetOrganizer().Address = participant.Email;
            //        //slave.GetOrganizer().Name = participant.Email;
            //        //Workaround: Assign organizer at least as first participant and as sent on behalf
            //        Outlook.Recipient recipient = slave.Recipients.Add(participant.Email);
            //        recipient.Type = (int)Outlook.OlMeetingRecipientType.olOrganizer; //Doesn't work (is ignored):
            //        if (recipient.Resolve())
            //        {

            //            const string PR_SENT_ON_BEHALF = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"; //-->works, but only on behalf, not organizer
            //            //const string PR_SENT_REPRESENTING_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x00410102";
            //            //const string PR_SENDER_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F";//-->Doesn't work: ComException, operation failed
            //            //const string PR_SENDER_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0C190102";//-->Doesn't work: ComException, operation failed
            //            //const string PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"; //-->Doesn't work: ComException, operation failed
            //            //const string PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F";//-->Doesn't work: ComException, operation failed
                      
            //            Microsoft.Office.Interop.Outlook.PropertyAccessor accessor = slave.PropertyAccessor;
            //            accessor.SetProperty(PR_SENT_ON_BEHALF, participant.Email);

            //            //const string PR_RECIPIENT_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x5FFD0003"; //-->Doesn't work: UnauthorizedAccessException, operation not allowed
            //            //Microsoft.Office.Interop.Outlook.PropertyAccessor accessor = recipient.PropertyAccessor;
            //            //accessor.SetProperty(PR_RECIPIENT_FLAGS, 3);
            //            //object test = accessor.GetProperty(PR_RECIPIENT_FLAGS);
            //        }

            //        break; //One Organizer is enough
            //    }

            //}

            ////Add remaining particpants
            //foreach (EventAttendee participant in master.Participants)
            //{
            //    if (participant.Rel != EventAttendee.RelType.EVENT_ORGANIZER && participant.Email != Syncronizer.UserName)
            //    {
            //        Outlook.Recipient recipient = slave.Recipients.Add(participant.Email);
            //        recipient.Resolve();

            //        //ToDo: Doesn't work because MeetingResponseStatus is readonly, maybe use PropertyAccessor?
            //        //switch (participant.Attendee_Status.Value)
            //        //{
            //        //    case Google.GData.Extensions.EventAttendee.AttendeeStatus.EVENT_ACCEPTED: recipient.MeetingResponseStatus = (int)Outlook.OlMeetingResponse.olMeetingAccepted; break;
            //        //    case Google.GData.Extensions.EventAttendee.AttendeeStatus.EVENT_DECLINED: recipient.MeetingResponseStatus = (int)Outlook.OlMeetingResponse.olMeetingDeclined; break;
            //        //    case Google.GData.Extensions.EventAttendee.AttendeeStatus.EVENT_TENTATIVE: recipient.MeetingResponseStatus = (int)Outlook.OlMeetingResponse.olMeetingTentative;
            //        //}
            //        if (participant.Attendee_Type != null)
            //        {
            //            switch (participant.Attendee_Type.Value)
            //            {
            //                case Google.GData.Extensions.EventAttendee.AttendeeType.EVENT_OPTIONAL: recipient.Type = (int)Outlook.OlMeetingRecipientType.olOptional; break;
            //                case Google.GData.Extensions.EventAttendee.AttendeeType.EVENT_REQUIRED: recipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired; break;
            //            }
            //        }

            //    }

            //}
            //End ToDo
            #endregion


            //slave.RequiredAttendees = master.RequiredAttendees;

            //slave.OptionalAttendees = master.OptionalAttendees;
            //slave.Resources = master.Resources;


            slave.ReminderSet = false;
            if (master.Reminders.UseDefault != null)
                slave.ReminderSet = master.Reminders.UseDefault.Value;
            if (master.Reminders.Overrides != null)
                foreach (var reminder in master.Reminders.Overrides)
                {
                    if (reminder.Method == "popup" && reminder.Minutes != null)
                    {
                        slave.ReminderSet = true;
                        slave.ReminderMinutesBeforeStart = reminder.Minutes.Value;
                    }
                }                        
            
            UpdateRecurrence(master, slave);

            //Sensivity update is only allowed for single appointments or the master
            if (!slave.IsRecurring || slave.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster) 
            {
                switch (master.Visibility)
                {
                    case "confidential": //ToDo, currently not supported by Google Web App GUI and Outlook 2010: slave.Sensitivity = Outlook.OlSensitivity.olConfidential; break;               
                    case "private": slave.Sensitivity = Outlook.OlSensitivity.olPrivate; break;
                    default: slave.Sensitivity = Outlook.OlSensitivity.olNormal; break;
                }
            }
    
        }

        public static void UpdateRecurrence(Outlook.AppointmentItem master, Event slave)
        {
            try
            {                               
                
                if (!master.IsRecurring)
                {
                    if (slave.Recurrence != null)
                        slave.Recurrence = null;
                    return;
                }

                if (master.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster)
                    return;

                Outlook.RecurrencePattern masterRecurrence = master.GetRecurrencePattern();

                string slaveRecurrence = string.Empty;

                //string format = dateFormat;
                //string key = VALUE + "=" + DATE;
                //if (!master.AllDayEvent)
                //{
                //    format += "'T'"+timeFormat;
                //    key = VALUE + "=" + DATETIME;
                //}

                ////For Debugging only:
                ////if (master.Subject == "IFX_CMF-Alignment - [Conference Number: 44246/ Password: 37757]")
                ////     throw new Exception ("Debugging: IFX_CMF-Alignment - [Conference Number: 44246/ Password: 37757]");

                ////ToDo: Find a way how to handle timezones, per default GMT (UTC+0:00) is taken
                ////if (master.StartTimeZone.ID == "W. Europe Standard Time")                
                ////    key = TZID + "=" + "Europe/Berlin";
                ////else if (master.StartTimeZone.ID == "Singapore Standard Time")
                ////    key = TZID + "=" + "Asia/Singapore";
                //if (!string.IsNullOrEmpty(Syncronizer.Timezone))
                //    key = TZID + "=" + Syncronizer.Timezone;

                //DateTime date = masterRecurrence.PatternStartDate.Date;
                ////DateTime time = new DateTime(date.Year, date.Month, date.Day, masterRecurrence.StartTime.Hour, masterRecurrence.StartTime.Minute, masterRecurrence.StartTime.Second);
                //DateTime time = new DateTime(date.Year, date.Month, date.Day, master.Start.Hour, master.Start.Minute, master.Start.Second);
                

                // The recurrence element contains various values that
                // are not used in Calendar v3 such as DTSTART
                // and DTEND
                // The recurrence element contains a list of string
                // representing an RFC 2445 RRULE, EXRULE, RDATE
                // or EXDATE rule.

                //slaveRecurrence += DTSTART;                    
                //slaveRecurrence += ";" + key + ":" + time.ToString(format) + "\r\n";

                //time = time.AddMinutes(masterRecurrence.Duration);             
                ////time = new DateTime(date.Year, date.Month, date.Day, masterRecurrence.EndTime.Hour, masterRecurrence.EndTime.Minute, masterRecurrence.EndTime.Second);               
                
                //slaveRecurrence += DTEND;
                //slaveRecurrence += ";"+key+":" + time.ToString(format) + "\r\n";

                if (slave.Recurrence == null)
                    slave.Recurrence = new List<string>();
                else
                    slave.Recurrence.Clear();
                slaveRecurrence = RRULE + ":" + FREQ +"=";
                switch (masterRecurrence.RecurrenceType)
                {
                    case Outlook.OlRecurrenceType.olRecursDaily: slaveRecurrence += DAILY; break;
                    case Outlook.OlRecurrenceType.olRecursWeekly: slaveRecurrence += WEEKLY; break;
                    case Outlook.OlRecurrenceType.olRecursMonthly: 
                    case Outlook.OlRecurrenceType.olRecursMonthNth: slaveRecurrence += MONTHLY; break;
                    case Outlook.OlRecurrenceType.olRecursYearly:
                    case Outlook.OlRecurrenceType.olRecursYearNth: slaveRecurrence += YEARLY; break;
                    default: throw new NotSupportedException("RecurrenceType not supported by Google: " + masterRecurrence.RecurrenceType);                                     
                }

                string byDay = string.Empty;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olMonday) == Outlook.OlDaysOfWeek.olMonday)
                    byDay = MO;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olTuesday) == Outlook.OlDaysOfWeek.olTuesday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TU;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olWednesday) == Outlook.OlDaysOfWeek.olWednesday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + WE;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olThursday) == Outlook.OlDaysOfWeek.olThursday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TH;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olFriday) == Outlook.OlDaysOfWeek.olFriday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + FR;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olSaturday) == Outlook.OlDaysOfWeek.olSaturday)
                   byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SA;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olSunday) == Outlook.OlDaysOfWeek.olSunday)
                   byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SU;

                if (!string.IsNullOrEmpty(byDay))
                {
                    if (masterRecurrence.Instance != 0)
                    {
                        if (masterRecurrence.Instance >= 1 && masterRecurrence.Instance <= 4)
                            byDay = masterRecurrence.Instance + byDay;
                        else if (masterRecurrence.Instance == 5)
                            slaveRecurrence += ";" + BYSETPOS + "=-1";
                        else
                            throw new NotSupportedException("Outlook Appointment Instances 1-4 and 5 (last) are allowed but was: " + masterRecurrence.Instance);
                    }
                    slaveRecurrence += ";" + BYDAY + "=" + byDay;
                }

                if (masterRecurrence.DayOfMonth != 0)
                    slaveRecurrence += ";" + BYMONTHDAY + "=" + masterRecurrence.DayOfMonth;

                if (masterRecurrence.MonthOfYear != 0)
                    slaveRecurrence += ";" + BYMONTH + "=" + masterRecurrence.MonthOfYear;

                if (masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly && 
                    masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth &&
                    masterRecurrence.Interval > 1 ||
                    masterRecurrence.Interval > 12)
                {
                    if (masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly &&
                        masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth)
                        slaveRecurrence += ";" + INTERVAL+ "=" + masterRecurrence.Interval;
                    else
                        slaveRecurrence += ";" + INTERVAL + "=" + masterRecurrence.Interval/12;
                }
                
                //format = dateFormat;
                if (masterRecurrence.PatternEndDate.Date != outlookDateMin &&
                    masterRecurrence.PatternEndDate.Date != outlookDateMax)
                {
                    slaveRecurrence += ";" + UNTIL + "=" + masterRecurrence.PatternEndDate.Date.AddDays(master.AllDayEvent?0:1).ToString(dateFormat);
                }
                //else if (masterRecurrence.Occurrences > 0)
                //{
                //    slaveRecurrence += ";" + COUNT + "=" + masterRecurrence.Occurrences;
                //}

                slave.Recurrence.Add(slaveRecurrence);
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }


        }

        /// <summary>
        /// Update Recurrence pattern from Google by parsing the string, see also specification http://tools.ietf.org/html/rfc2445
        /// </summary>
        /// <param name="master"></param>
        /// <param name="slave"></param>
        public static void UpdateRecurrence(Event master, Outlook.AppointmentItem slave)
        {
            var masterRecurrence = master.Recurrence;   
            if (masterRecurrence == null)
            {
                if (slave.IsRecurring && slave.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                    slave.ClearRecurrencePattern();

                return;
            }            


            try
            {  
                 
                 Outlook.RecurrencePattern slaveRecurrence = slave.GetRecurrencePattern();

                 if (master.Start != null && !string.IsNullOrEmpty(master.Start.Date))
                 {
                     slaveRecurrence.StartTime = DateTime.Parse(master.Start.Date);
                     slaveRecurrence.PatternStartDate = DateTime.Parse(master.Start.Date);
                 }
                 else if (master.Start != null && master.Start.DateTime != null)
                 {
                     slaveRecurrence.StartTime = master.Start.DateTime.Value;
                     slaveRecurrence.PatternStartDate = master.Start.DateTime.Value;
                 }                 

                 
                //string[] patterns = masterRecurrence.Value.Split(new char[] {'\r','\n'}, StringSplitOptions.RemoveEmptyEntries);
                //foreach (string pattern in patterns)
                //{
                //    if (pattern.StartsWith(DTSTART)) 
                //    {
                //        //DTSTART;VALUE=DATE:20070501
                //        //DTSTART;TZID=US-Eastern:19970905T090000
                //        string[] parts = pattern.Split(new char[] {';',':'});
                        
                //        slaveRecurrence.StartTime = GetDateTime(parts[parts.Length-1]);
                //        slaveRecurrence.PatternStartDate = GetDateTime(parts[parts.Length - 1]);
                //        break;
                //    }
                //}

                if (master.End != null && !string.IsNullOrEmpty(master.End.Date))
                     slaveRecurrence.EndTime = DateTime.Parse(master.End.Date);
                 if (master.End != null && master.End.DateTime != null)
                     slaveRecurrence.EndTime = master.End.DateTime.Value;
                 
                //foreach (string pattern in patterns)
                //{
                //    if (pattern.StartsWith(DTEND))
                //    {
                //        string[] parts = pattern.Split(new char[] { ';', ':' });
                        
                //        slaveRecurrence.EndTime = GetDateTime(parts[parts.Length-1]);
                //        //Don't update, otherwise it will end after first occurrence: slaveRecurrence.PatternEndDate = GetDateTime(parts[parts.Length - 1]);
                        
                //        break;
                //    }
                //}

                foreach (string pattern in master.Recurrence)
                {
                    if (pattern.StartsWith(RRULE))
                    {
                        string[] parts = pattern.Split(new char[] { ';', ':' });

                        int instance = 0;
                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                string[] days = part.Split(',');
                                foreach (string day in days)
                                {
                                    string dayValue = day.Substring(day.IndexOf("=") + 1);
                                    if (dayValue.StartsWith("1"))
                                        instance = 1;
                                    else if (dayValue.StartsWith("2"))
                                        instance = 2;
                                    else if (dayValue.StartsWith("3"))
                                        instance = 3;
                                    else if (dayValue.StartsWith("4"))
                                        instance = 4;


                                    break;
                                }

                                break;
                            }

                        }

                        foreach (string part in parts)
                        {

                            if (part.StartsWith(BYSETPOS))
                            {
                                string pos = part.Substring(part.IndexOf("=") + 1);

                                if (pos.Trim() == "-1")
                                    instance = 5;
                                else
                                   throw new NotSupportedException("Only 'BYSETPOS=-1' is allowed by Outlook, but it was: " + part);                                    

                              break;
                            }
                        }


                        foreach (string part in parts)
                        {
                            if (part.StartsWith(FREQ))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case DAILY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily; break;
                                    case WEEKLY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly; break;
                                    case MONTHLY:
                                        if (instance == 0)
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly;
                                        else
                                        {
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthNth;
                                            slaveRecurrence.Instance = instance;
                                        }
                                        break;
                                    case YEARLY:
                                        if (instance == 0)
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
                                        else
                                        {
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearNth;
                                            slaveRecurrence.Instance = instance;
                                        }
                                        break;
                                    default: throw new NotSupportedException("RecurrenceType not supported by Outlook: " + part);
                                                                        
                                }
                                break;
                            }

                        }

                        foreach (string part in parts)
                        {

                            if (part.StartsWith(BYDAY))
                            {
                                Outlook.OlDaysOfWeek dayOfWeek = slaveRecurrence.DayOfWeekMask;
                                string[] days = part.Split(',');
                                foreach (string day in days)
                                {
                                    string dayValue = day.Substring(day.IndexOf("=") + 1);                                    

                                    switch (dayValue.Trim(new char[] { '1', '2', '3', '4', ' ' }))
                                    {
                                        case MO: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olMonday; break;
                                        case TU: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olTuesday; break;
                                        case WE: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olWednesday; break;
                                        case TH: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olThursday; break;
                                        case FR: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olFriday; break;
                                        case SA: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olSaturday; break;
                                        case SU: dayOfWeek = dayOfWeek | Outlook.OlDaysOfWeek.olSunday; break;

                                    }
                                    //Don't break because multiple days possible;
                                }

                                if (slaveRecurrence.DayOfWeekMask != dayOfWeek && dayOfWeek != 0)
                                    slaveRecurrence.DayOfWeekMask = dayOfWeek;

                                break;
                            }
                        }


                        foreach (string part in parts)
                        {
                            if (part.StartsWith(INTERVAL))
                            {
                                int interval = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                if (slaveRecurrence.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearly ||
                                    slaveRecurrence.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearNth)
                                {
                                    interval = interval * 12; // must be expressed in months
                                }
                                slaveRecurrence.Interval = interval;
                                break;
                            }

                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(COUNT))
                            {                                
                                slaveRecurrence.Occurrences = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                            else if (part.StartsWith(UNTIL))
                            {
                                //either UNTIL or COUNT may appear in a 'recur',
                                //but UNTIL and COUNT MUST NOT occur in the same 'recur'
                                slaveRecurrence.PatternEndDate = GetDateTime(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }

                        





                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYMONTHDAY))
                            {
                                slaveRecurrence.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYMONTH + "="))
                            {
                                slaveRecurrence.MonthOfYear = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }


                        break;
                    }

                }
                

            }
            catch (Exception ex)
            {
              //ErrorHandler.Handle(ex);
              Logger.Log("Error updating event's BYSETPOS: " + master.Summary + " - " + Synchronizer.GetTime(master) + ": " + ex.Message, EventType.Error);
            }


        }

        public static bool UpdateRecurrenceExceptions(Outlook.AppointmentItem master, Event slave, Synchronizer sync)
        {
            
            bool ret = false;

            Outlook.Exceptions exceptions = master.GetRecurrencePattern().Exceptions;

            if (exceptions != null && exceptions.Count != 0)                
            {
                foreach (Outlook.Exception exception in exceptions)
                {
                    if (!exception.Deleted) 
                    {
                        //Add exception time (but only if in given time range
                        if ((Synchronizer.MonthsInPast == 0 || exception.AppointmentItem.End >= DateTime.Now.AddMonths(-Synchronizer.MonthsInPast)) &&
                             (Synchronizer.MonthsInFuture == 0 || exception.AppointmentItem.Start <= DateTime.Now.AddMonths(Synchronizer.MonthsInFuture)))
                        {
                            //slave.Times.Add(new Google.GData.Extensions.When(exception.AppointmentItem.Start, exception.AppointmentItem.Start, exception.AppointmentItem.AllDayEvent));
                            var googleRecurrenceException = Factory.NewEvent();
                            //if (slave.Sequence != null)
                            //    googleRecurrenceException.Sequence = slave.Sequence + 1;
                            googleRecurrenceException.RecurringEventId = slave.Id;
                            //googleRecurrenceException.OriginalEvent.Href = ??? 
                            googleRecurrenceException.OriginalStartTime = new EventDateTime();
                            if (master.AllDayEvent == true)
                                googleRecurrenceException.OriginalStartTime.Date = exception.OriginalDate.ToString("yyyy-MM-dd");
                            else
                            {
                                //DateTime start = exception.OriginalDate.AddHours(master.Start.Hour).AddMinutes(master.Start.Minute).AddSeconds(master.Start.Second);
                                googleRecurrenceException.OriginalStartTime.DateTime = exception.OriginalDate;
                            }


                            try
                            {
                                sync.UpdateAppointment(exception.AppointmentItem, ref googleRecurrenceException);                               
                                //googleRecurrenceExceptions.Add(googleRecurrenceException);                            

                                ret = true;
                            }
                            catch (Exception ex)
                            {
                                //should not happen
                                Logger.Log(ex.Message, EventType.Error);
                            }
                        }
                    }
                    else
                    {//Delete exception time

                    //    //for (int i=slave.Times.Count;i>0;i--)
                    //    //{
                    //    //    When time = slave.Times[i-1];
                    //    //    if (time.StartTime.Equals(exception.AppointmentItem.Start))
                    //    //    {
                    //    //        slave.Times.Remove(time);
                    //    //        ret = true;
                    //    //        break;
                    //    //    }
                    //    //}

                    //    //for (int i = googleRecurrenceExceptions.Count; i > 0;i-- )
                    //    //{
                    //    //    if (googleRecurrenceExceptions[i-1].Times[0].StartTime.Equals(exception.OriginalDate))
                    //    //    {
                    //    //        googleRecurrenceExceptions[i - 1].Delete();
                    //    //        googleRecurrenceExceptions.RemoveAt(i - 1);
                    //    //    }
                    //    //}

                    //    //ToDo: Doesn't work for all recurrences                        
                    //    var googleRecurrenceException = sync.GetGoogleAppointmentByStartDate(slave.Id, exception.OriginalDate);                                 

                    //    if (googleRecurrenceException != null)
                    //        googleRecurrenceException.Delete();
                        
                        if ((Synchronizer.MonthsInPast == 0 || exception.OriginalDate >= DateTime.Now.AddMonths(-Synchronizer.MonthsInPast)) &&
                             (Synchronizer.MonthsInFuture == 0 || exception.OriginalDate <= DateTime.Now.AddMonths(Synchronizer.MonthsInFuture)))
                        {
                            //First create deleted occurrences, to delete it later again
                            var googleRecurrenceException = Factory.NewEvent();
                            //if (slave.Sequence != null)
                            //    googleRecurrenceException.Sequence = slave.Sequence + 1;
                            googleRecurrenceException.RecurringEventId = slave.Id;
                            //googleRecurrenceException.OriginalEvent.Href = ???
                            DateTime start = exception.OriginalDate.AddHours(master.Start.Hour).AddMinutes(master.Start.Minute).AddSeconds(master.Start.Second);
                            googleRecurrenceException.OriginalStartTime =  new EventDateTime();
                            googleRecurrenceException.OriginalStartTime.TimeZone = slave.Start.TimeZone;
                            
                            if (master.AllDayEvent)
                            {
                                googleRecurrenceException.OriginalStartTime.Date = start.ToString("yyyy-MM-dd");
                                googleRecurrenceException.End.Date = start.AddMinutes(master.Duration).ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                googleRecurrenceException.OriginalStartTime.DateTime = start;
                                googleRecurrenceException.End.DateTime = start.AddMinutes(master.Duration);
                            }
                            googleRecurrenceException.Start = googleRecurrenceException.OriginalStartTime;
                           
                            googleRecurrenceException.Summary = master.Subject;

                            try
                            {
                                googleRecurrenceException = sync.SaveGoogleAppointment(googleRecurrenceException);
                                //googleRecurrenceExceptions.Add(googleRecurrenceException);                                  

                                //ToDo: check promptDeletion and syncDeletion options
                                sync.EventRequest.Delete(Synchronizer.SyncAppointmentsGoogleFolder, googleRecurrenceException.Id).Execute();
                                Logger.Log("Deleted obsolete recurrence exception from Google: " + master.Subject + " - " + exception.OriginalDate, EventType.Information);
                                //sync.DeletedCount++;

                                ret = true;
                            }
                            catch (Exception ex)
                            {
                                //usually only an error is thrown, if an already cancelled event is to be deleted again
                                Logger.Log(ex.Message, EventType.Debug);
                            }


                        }
                        

                    }
                }
            }

            return ret;
        }

        public static bool UpdateRecurrenceExceptions(List<Event> googleRecurrenceExceptions, Outlook.AppointmentItem slave, Synchronizer sync)
        {
            bool ret = false;

            for (int i = 0; i < googleRecurrenceExceptions.Count; i++)
            {
                Event googleRecurrenceException = googleRecurrenceExceptions[i];
                //if (slave == null || !slave.IsRecurring || slave.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster)
                //    Logger.Log("Google Appointment with OriginalEvent found, but Outlook is not recurring: " + googleAppointment.Summary + " - " + GetTime(googleAppointment), EventType.Warning);
                //else
                //{                         
                Outlook.AppointmentItem outlookRecurrenceException = null;
                try
                {
                    var slaveRecurrence = slave.GetRecurrencePattern();
                    if (googleRecurrenceException.OriginalStartTime != null && !string.IsNullOrEmpty(googleRecurrenceException.OriginalStartTime.Date))
                        outlookRecurrenceException = slaveRecurrence.GetOccurrence(DateTime.Parse(googleRecurrenceException.OriginalStartTime.Date));
                    else if (googleRecurrenceException.OriginalStartTime != null && googleRecurrenceException.OriginalStartTime.DateTime != null)
                        outlookRecurrenceException = slaveRecurrence.GetOccurrence(googleRecurrenceException.OriginalStartTime.DateTime.Value);
                }
                catch (Exception ignored)
                {
                    Logger.Log("Google Appointment with OriginalEvent found, but Outlook occurrence not found: " + googleRecurrenceException.Summary + " - " + googleRecurrenceException.OriginalStartTime.DateTime + ": " + ignored, EventType.Debug);
                }                


                if (outlookRecurrenceException != null)
                {
                    //myInstance.Subject = googleAppointment.Summary;
                    //myInstance.Start = googleAppointment.Times[0].StartTime;
                    //myInstance.End = googleAppointment.Times[0].EndTime;
                    DateTime? timeMin = null;
                    DateTime? timeMax = null;
                    if (googleRecurrenceException.Start != null && !string.IsNullOrEmpty(googleRecurrenceException.Start.Date))
                        timeMin = DateTime.Parse(googleRecurrenceException.Start.Date);
                    else if (googleRecurrenceException.Start != null)
                        timeMin = googleRecurrenceException.Start.DateTime;

                    if (googleRecurrenceException.End != null && !string.IsNullOrEmpty(googleRecurrenceException.End.Date))
                        timeMax = DateTime.Parse(googleRecurrenceException.End.Date);
                    else if (googleRecurrenceException.End != null)
                        timeMax = googleRecurrenceException.End.DateTime;

                    googleRecurrenceException = sync.LoadGoogleAppointments(googleRecurrenceException.Id, 0, 0, timeMin, timeMax); //Reload, just in case it was updated by master recurrence                                
                    if (googleRecurrenceException != null)
                    {
                        if (googleRecurrenceException.Status.Equals("cancelled"))
                        {
                            outlookRecurrenceException.Delete();
                            string timeToLog = null;
                            if (googleRecurrenceException.OriginalStartTime != null)
                            {   
                                timeToLog = googleRecurrenceException.OriginalStartTime.Date;
                                if (string.IsNullOrEmpty(timeToLog) && googleRecurrenceException.OriginalStartTime.DateTime != null)
                                timeToLog =  googleRecurrenceException.OriginalStartTime.DateTime.Value.ToString();
                            }

                            Logger.Log("Deleted obsolete recurrence exception from Outlook: " + slave.Subject + " - " + timeToLog, EventType.Information);
                        }
                        else
                        {
                            if (sync.UpdateAppointment(ref googleRecurrenceException, outlookRecurrenceException, null))
                            {
                                outlookRecurrenceException.Save();
                                Logger.Log("Updated recurrence exception from Google to Outlook: " + googleRecurrenceException.Summary + " - " + Synchronizer.GetTime(googleRecurrenceException), EventType.Information);
                            }
                        }
                        ret = true;

                        
                    }
                    else
                        Logger.Log("Error updating recurrence exception from Google to Outlook (couldn't be reload from Google): " + outlookRecurrenceException.Subject + " - " + outlookRecurrenceException.Start, EventType.Information);

                }
                

                //}

            }

            return ret;
        }

        private static DateTime GetDateTime(string dateTime)
        {
            string format = dateFormat;
            if (dateTime.Contains("T"))
                format += "'T'"+timeFormat;
            if (dateTime.EndsWith("Z"))
                format += "'Z'";
            return DateTime.ParseExact(dateTime, format, new System.Globalization.CultureInfo("en-US"));
        }

        //internal static EventAttendee GetOrganizer(Event googleAppointment)
        //{
        //    foreach (var person in googleAppointment.Participants)
        //    {

        //        if (person.Rel == EventAttendee.RelType.EVENT_ORGANIZER)
        //        {
        //            return person;
        //        }
        //    }
        //    return null;
        //}

        //internal static bool IsOrganizer(EventAttendee person)
        //{
        //    if (person != null && person.Email != null && person.Email.Trim().Equals(Syncronizer.UserName.Trim(), StringComparison.InvariantCultureIgnoreCase))
        //        return true;
        //    else
        //        return false;
        //}

        internal static bool IsOrganizer(string email)
        {
            string userName = Synchronizer.UserName.Trim().ToLower().Replace("@googlemail.", "@gmail.");
            if (email != null)
            {
                email = email.Trim().ToLower().Replace("@googlemail.", "@gmail.");
                if (email.Equals(userName, StringComparison.InvariantCultureIgnoreCase))
                    return true;
                else
                    return false;
            }
            return false;
        }

        //internal static string GetOrganizer(Outlook.AppointmentItem outlookAppointment)
        //{
        //    Outlook.AddressEntry organizer = outlookAppointment.GetOrganizer();
        //    if (organizer != null)
        //    {
        //        if (string.IsNullOrEmpty(organizer.Address))
        //            return organizer.Address;
        //        else
        //            return organizer.Name;
        //    }

        //    return outlookAppointment.Organizer;            


        //}

        //internal static bool IsOrganizer(string person, Outlook.AppointmentItem outlookAppointment)
        //{
        //    if (!string.IsNullOrEmpty(person) && 
        //        (person.Trim().Equals(outlookAppointment.Session.CurrentUser.Address, StringComparison.InvariantCultureIgnoreCase) || 
        //        person.Trim().Equals(outlookAppointment.Session.CurrentUser.Name, StringComparison.InvariantCultureIgnoreCase)
        //        ))
        //        return true;
        //    else
        //        return false;
        //}
        

    }
}
