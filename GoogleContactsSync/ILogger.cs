using Google.Contacts;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace GoContactSyncMod
{
    enum EventType
    {
        Debug,
        Information,
        Warning,
        Error
    }

    struct LogEntry
    {
        public DateTime date;
        public EventType type;
        public string msg;

        public LogEntry(DateTime _date, EventType _type, string _msg)
        {
            date = _date; type = _type; msg = _msg;
        }
    }

    static class Logger
    {
        public static List<LogEntry> messages = new List<LogEntry>();
        public delegate void LogUpdatedHandler(string Message);
        public static event LogUpdatedHandler LogUpdated;
        private static StreamWriter logwriter;

        public static readonly string Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
        public static readonly string AuthFolder = Folder + "\\Auth\\";

        static Logger()
        {
            try
            {
                if (!Directory.Exists(Folder))
                    Directory.CreateDirectory(Folder);

                if (!Directory.Exists(AuthFolder))
                    Directory.CreateDirectory(AuthFolder);

                string logFileName = Folder + "log.txt";

                //If log file is bigger than 1 MB, move it to backup file and create new file
                FileInfo logFile = new FileInfo(logFileName);
                if (logFile.Exists && logFile.Length >= 1000000)
                    File.Move(logFileName, logFileName + "_" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss"));

                logwriter = new StreamWriter(logFileName, true);

                logwriter.WriteLine("[Start Rolling]");
                logwriter.Flush();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }

        public static void Close()
        {
            try
            {
                if (logwriter != null)
                {
                    logwriter.WriteLine("[End Rolling]");
                    logwriter.Flush();
                    logwriter.Close();
                }
            }
            catch (Exception e)
            {
                ErrorHandler.Handle(e);
            }
        }

        private static string formatMessage(string message, EventType eventType)
        {
            return string.Format("{0}:{1}{2}", eventType, Environment.NewLine, message);
        }

        private static string GetLogLine(LogEntry entry)
        {
            return string.Format("[{0} | {1}]\t{2}\r\n", entry.date, entry.type, entry.msg);
        }

        public static void Log(string message, EventType eventType)
        {
            LogEntry new_logEntry = new LogEntry(DateTime.Now, eventType, message);
            messages.Add(new_logEntry);

            try
            {
                logwriter.Write(GetLogLine(new_logEntry));
                logwriter.Flush();
            }
            catch (Exception)
            {
                //ignore it, because if you handle this error, the handler will again log the message
                //ErrorHandler.Handle(ex);
            }

            //Populate LogMessage to all subscribed Logger-Outputs, but only if not Debug message, Debug messages are only logged to logfile
            if (LogUpdated != null && eventType > EventType.Debug)
                LogUpdated(GetLogLine(new_logEntry));
        }

        public static void Log(Exception ex, EventType eventType)
        {
            CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            if (ex.InnerException != null)
            {
                Log("Inner Exception Type: " + ex.InnerException.GetType().ToString(), eventType);

                COMException ci = ex.InnerException as COMException;
                if (ci != null)
                {
                    Log("Inner Error Code: " + ci.ErrorCode.ToString("X"), eventType);
                }
                Log("Inner Exception: " + ex.InnerException.Message, eventType);
                Log("Inner Source: " + ex.InnerException.Source, eventType);
                if (ex.InnerException.StackTrace != null)
                {
                    Log("Inner Stack Trace: " + ex.InnerException.StackTrace, eventType);
                }
            }
            Log("Exception Type: " + ex.GetType().ToString(), eventType);
            COMException c = ex as COMException;
            if (c != null)
            {
                Log("Error Code: " + c.ErrorCode.ToString("X"), eventType);
            }
            Log("Exception: " + ex.Message, eventType);
            Log("Source: " + ex.Source, eventType);
            if (ex.StackTrace != null)
            {
                Log("Stack Trace: " + ex.StackTrace, eventType);
            }

            Thread.CurrentThread.CurrentCulture = oldCI;
            Thread.CurrentThread.CurrentUICulture = oldCI;
        }

        public static void Log(Google.Apis.Calendar.v3.Data.Event e, EventType eventType)
        {
            Log("*** Google event ***", eventType);
            Log(" - AnyoneCanAddSelf: " + (e.AnyoneCanAddSelf != null ? e.AnyoneCanAddSelf.ToString() : "null"), eventType);
            if (e.Attachments != null)
            {
                Log(" - Attachments:", eventType);
                foreach (var a in e.Attachments)
                {
                    Log("  - Title: " + (a.Title != null ? a.Title : "null"), eventType);
                }
            }
            if (e.Attendees != null)
            {
                Log(" - Attendees:", eventType);
                foreach (var a in e.Attendees)
                {
                    Log("  - DisplayName: " + (a.DisplayName != null ? a.DisplayName : "null"), eventType);
                }
            }
            Log(" - AttendeesOmitted: " + (e.AttendeesOmitted != null ? e.AttendeesOmitted.ToString() : "null"), eventType);
            Log(" - ColorId: " + (e.ColorId != null ? e.ColorId : "null"), eventType);
            Log(" - Created: " + (e.Created != null ? e.Created.ToString() : "null"), eventType);
            Log(" - CreatedRaw: " + (e.CreatedRaw != null ? e.CreatedRaw : "null"), eventType);
            if (e.Creator != null)
            {
                Log(" - Creator:", eventType);
                Log("  - DisplayName: " + (e.Creator.DisplayName != null ? e.Creator.DisplayName : "null"), eventType);
            }
            Log(" - Description: " + (e.Description != null ? e.Description : "null"), eventType);
            if (e.End != null)
            {
                Log(" - End:", eventType);
                if (!string.IsNullOrEmpty(e.End.Date))
                    Log("  - Date: " + e.End.Date, eventType);
                if (e.End.DateTime != null)
                    Log("  - DateTime: " + e.End.DateTime.Value.ToString(), eventType);
                if (!string.IsNullOrEmpty(e.End.TimeZone))
                    Log("  - TimeZone: " + e.End.TimeZone, eventType);
            }
            Log(" - EndTimeUnspecified: " + (e.EndTimeUnspecified != null ? e.EndTimeUnspecified.ToString() : "null"), eventType);
            if (e.ExtendedProperties != null)
            {
                Log(" - ExtendedProperties:", eventType);
                if (e.ExtendedProperties.Shared != null)
                {
                    Log("  - Shared:", eventType);
                    foreach (var p in e.ExtendedProperties.Shared)
                    {
                        Log("   - Key: " + (p.Key != null ? p.Key : "null"), eventType);
                        Log("   - Value: " + (p.Value != null ? p.Value : "null"), eventType);
                    }
                }
                if (e.ExtendedProperties.Private__ != null)
                {
                    Log("  - Private__:", eventType);
                    foreach (var p in e.ExtendedProperties.Private__)
                    {
                        Log("   - Key: " + (p.Key != null ? p.Key : "null"), eventType);
                        Log("   - Value: " + (p.Value != null ? p.Value : "null"), eventType);
                    }
                }
            }
            if (e.Gadget != null)
            {
                Log(" - Gadget:", eventType);
                if (!string.IsNullOrEmpty(e.Gadget.Title))
                    Log("  - Title: " + e.Gadget.Title, eventType);
            }
            Log(" - GuestsCanInviteOthers: " + (e.GuestsCanInviteOthers != null ? e.GuestsCanInviteOthers.ToString() : "null"), eventType);
            Log(" - GuestsCanModify: " + (e.GuestsCanModify != null ? e.GuestsCanModify.ToString() : "null"), eventType);
            Log(" - GuestsCanSeeOtherGuests: " + (e.GuestsCanSeeOtherGuests != null ? e.GuestsCanSeeOtherGuests.ToString() : "null"), eventType);
            Log(" - HangoutLink: " + (e.HangoutLink != null ? e.HangoutLink : "null"), eventType);
            Log(" - HtmlLink: " + (e.HtmlLink != null ? e.HtmlLink : "null"), eventType);
            Log(" - ICalUID: " + (e.ICalUID != null ? e.ICalUID : "null"), eventType);
            Log(" - Id: " + (e.Id != null ? e.Id : "null"), eventType);
            Log(" - Kind: " + (e.Kind != null ? e.Kind : "null"), eventType);
            Log(" - Location: " + (e.Location != null ? e.Location : "null"), eventType);
            Log(" - Locked: " + (e.Locked != null ? e.Locked.ToString() : "null"), eventType);
            if (e.Organizer != null)
            {
                Log(" - Organizer:", eventType);
                Log("  - DisplayName: " + (e.Organizer.DisplayName != null ? e.Organizer.DisplayName : "null"), eventType);
            }
            if (e.OriginalStartTime != null)
            {
                Log(" - OriginalStartTime:", eventType);
                if (!string.IsNullOrEmpty(e.OriginalStartTime.Date))
                    Log("  - Date: " + e.OriginalStartTime.Date, eventType);
                if (e.OriginalStartTime.DateTime != null)
                    Log("  - DateTime: " + e.OriginalStartTime.DateTime.Value.ToString(), eventType);
                if (!string.IsNullOrEmpty(e.OriginalStartTime.TimeZone))
                    Log("  - TimeZone: " + e.OriginalStartTime.TimeZone, eventType);
            }
            Log(" - PrivateCopy: " + (e.PrivateCopy != null ? e.PrivateCopy.ToString() : "null"), eventType);
            if (e.Recurrence != null)
            {
                Log(" - Recurrence:", eventType);
                foreach (var r in e.Recurrence)
                {
                    Log("  - : " + r, eventType);
                }
            }
            Log(" - RecurringEventId: " + (e.RecurringEventId != null ? e.RecurringEventId : "null"), eventType);
            if (e.Reminders != null)
            {
                Log(" - Reminders:", eventType);
                if (e.Reminders.UseDefault != null)
                    Log("  - UseDefault: " + e.Reminders.UseDefault.ToString(), eventType);
                if (e.Reminders.Overrides != null)
                {
                    Log("  - Overrides:", eventType);
                    foreach (var o in e.Reminders.Overrides)
                    {
                        Log("   - Minutes: " + (o.Minutes != null ? o.Minutes.ToString() : "null"), eventType);
                    }
                }
            }
            Log(" - Sequence: " + (e.Sequence != null ? e.Sequence.ToString() : "null"), eventType);
            if (e.Source != null)
            {
                Log(" - Source:", eventType);
                Log("  - Url: " + (e.Source.Url != null ? e.Source.Url : "null"), eventType);
            }
            if (e.Start != null)
            {
                Log(" - Start:", eventType);
                if (!string.IsNullOrEmpty(e.Start.Date))
                    Log("  - Date: " + e.Start.Date, eventType);
                if (e.Start.DateTime != null)
                    Log("  - DateTime: " + e.Start.DateTime.Value.ToString(), eventType);
                if (!string.IsNullOrEmpty(e.Start.TimeZone))
                    Log("  - TimeZone: " + e.Start.TimeZone, eventType);
            }
            Log(" - Status: " + (e.Status != null ? e.Status : "null"), eventType);
            Log(" - Summary: " + (e.Summary != null ? e.Summary : "null"), eventType);
            Log(" - Transparency: " + (e.Transparency != null ? e.Transparency : "null"), eventType);
            Log(" - Updated: " + (e.Updated != null ? e.Updated.ToString() : "null"), eventType);
            Log(" - UpdatedRaw: " + (e.UpdatedRaw != null ? e.UpdatedRaw : "null"), eventType);
            Log(" - Visibility: " + (e.Visibility != null ? e.Visibility : "null"), eventType);
            Log("*** Google event ***", eventType);
        }

        public static void Log(Contact c, EventType eventType)
        {
            Log("*** Google contact ***", eventType);
            if (c.AppControl != null)
            {
                Log(" - AppControl:", eventType);
            }
            if (c.AtomEntry != null)
            {
                Log(" - AtomEntry:", eventType);
            }
            Log(" - Author: " + (c.Author != null ? c.Author : "null"), eventType);
            if (c.BatchData != null)
            {
                Log(" - BatchData:", eventType);
                if (c.BatchData.Id != null)
                    Log("  - Id: " + c.BatchData.Id, eventType);
            }
            if (c.Categories != null)
            {
                Log(" - Categories:", eventType);
                foreach (var e in c.Categories)
                {
                    Log("  - Scheme: " + (e.Scheme != null ? e.Scheme : "null"), eventType);
                    Log("  - Term: " + (e.Term != null ? e.Term : "null"), eventType);
                }
            }
            if (c.ContactEntry != null)
            {
                Log(" - ContactEntry:", eventType);
                Log("  - Initials: " + (c.ContactEntry.Initials != null ? c.ContactEntry.Initials : "null"), eventType);
            }
            Log(" - Content: " + (c.Content != null ? c.Content : "null"), eventType);
            Log(" - Deleted: " + c.Deleted.ToString(), eventType);
            if (c.Emails != null)
            {
                Log(" - Emails:", eventType);
                foreach (var e in c.Emails)
                {
                    Log("  - Address: " + (e.Address != null ? e.Address : "null"), eventType);
                    Log("  - Label: " + (e.Label != null ? e.Label : "null"), eventType);
                    Log("  - Primary: " + e.Primary.ToString(), eventType);
                }
            }
            Log(" - ETag: " + (c.ETag != null ? c.ETag : "null"), eventType);
            if (c.ExtendedProperties != null)
            {
                Log(" - ExtendedProperties:", eventType);
                foreach (var e in c.ExtendedProperties)
                {
                    Log("  - Name: " + (e.Name != null ? e.Name : "null"), eventType);
                    Log("  - Value: " + (e.Value != null ? e.Value : "null"), eventType);
                }
            }
            if (c.GroupMembership != null)
            {
                Log(" - GroupMembership:", eventType);
                foreach (var e in c.GroupMembership)
                {
                    Log("  - HRef: " + (e.HRef != null ? e.HRef : "null"), eventType);
                }
            }
            Log(" - Id: " + (c.Id != null ? c.Id : "null"), eventType);
            if (c.IMs != null)
            {
                Log(" - IMs:", eventType);
                foreach (var e in c.IMs)
                {
                    Log("  - Value: " + (e.Value  != null ? e.Value : "null"), eventType);
                }
            }
            Log(" - IsDraft: " + c.IsDraft.ToString(), eventType);
            if (c.Languages != null)
            {
                Log(" - Languages:", eventType);
                foreach (var e in c.Languages)
                {
                    Log("  - Value: " + (e.Value != null ? e.Value : "null"), eventType);
                }
            }
            Log(" - Location: " + (c.Location != null ? c.Location : "null"), eventType);
            if (c.MediaSource != null)
            {
                Log(" - MediaSource:", eventType);
                Log("  - Name: " + (c.MediaSource.Name != null ? c.MediaSource.Name : "null"), eventType);
            }
            if (c.Name != null)
            {
                Log(" - Name:", eventType);
                Log("  - FamilyName: " + (c.Name.FamilyName != null ? c.Name.FamilyName : "null"), eventType);
                Log("  - FullName: " + (c.Name.FullName != null ? c.Name.FullName : "null"), eventType);
                Log("  - GivenName: " + (c.Name.GivenName != null ? c.Name.GivenName : "null"), eventType);
            }
            if (c.Organizations != null)
            {
                Log(" - Organizations:", eventType);
                foreach (var e in c.Organizations)
                {
                    Log("  - Name: " + (e.Name  != null ? e.Name : "null"), eventType);
                }
            }
            if (c.Phonenumbers != null)
            {
                Log(" - Phonenumbers:", eventType);
                foreach (var e in c.Phonenumbers)
                {
                    Log("  - Rel: " + (e.Rel != null ? e.Rel : "null"), eventType);
                    Log("  - Value: " + (e.Value != null ? e.Value : "null"), eventType);
                }
            }
            Log(" - PhotoEtag: " + (c.PhotoEtag != null ? c.PhotoEtag : "null"), eventType);
            if (c.PhotoUri != null)
            {
                Log(" - PhotoUri:", eventType);
                Log("  - OriginalString: " + (c.PhotoUri.OriginalString != null ? c.PhotoUri.OriginalString : "null"), eventType);
            } 
            if (c.PostalAddresses != null)
            {
                Log(" - PostalAddresses:", eventType);
                foreach (var e in c.PostalAddresses)
                {
                    Log("  - Stret: " + (e.Street != null ? e.Street : "null"), eventType);
                }
            }
            if (c.PrimaryEmail != null)
            {
                Log(" - PrimaryEmail:", eventType);
                Log("  - Value: " + (c.PrimaryEmail.Value != null ? c.PrimaryEmail.Value : "null"), eventType);
            }
            if (c.PrimaryIMAddress != null)
            {
                Log(" - PrimaryIMAddress:", eventType);
                Log("  - Value: " + (c.PrimaryIMAddress.Value != null ? c.PrimaryIMAddress.Value : "null"), eventType);
            }
            if (c.PrimaryPhonenumber != null)
            {
                Log(" - PrimaryPhonenumber:", eventType);
                Log("  - Value: " + (c.PrimaryPhonenumber.Value != null ? c.PrimaryPhonenumber.Value : "null"), eventType);
            }
            if (c.PrimaryPostalAddress != null)
            {
                Log(" - PrimaryPostalAddress:", eventType);
                Log("  - Street: " + (c.PrimaryPostalAddress.Street != null ? c.PrimaryPostalAddress.Street : "null"), eventType);
            }
            Log(" - ReadOnly: " + c.ReadOnly.ToString(), eventType);
            Log(" - Self: " + (c.Self != null ? c.Self : "null"), eventType);
            Log(" - Summary: " + (c.Summary != null ? c.Summary : "null"), eventType);
            Log(" - Title: " + (c.Title != null ? c.Title : "null"), eventType);
            Log(" - Updated: " + (c.Updated != null ? c.Updated.ToString() : "null"), eventType);
            Log("*** Google contact ***", eventType);
        }

        public static void ClearLog()
        {
            messages.Clear();
        }
    }
}