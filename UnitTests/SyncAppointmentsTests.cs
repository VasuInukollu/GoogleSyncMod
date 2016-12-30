using System;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using Google.Apis.Calendar.v3.Data;
using System.Runtime.InteropServices;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncAppointmentsTests
    {
        Synchronizer sync;

        static Logger.LogUpdatedHandler _logUpdateHandler = null;

        //Constants for test appointment
        const string name = "AN_OUTLOOK_TEST_APPOINTMENT";
        //readonly When whenDay = new When(DateTime.Now, DateTime.Now, true);
        //readonly When whenTime = new When(DateTime.Now, DateTime.Now.AddHours(1), false);
        //ToDo:const string groupName = "A TEST GROUP";

        [OneTimeSetUp]
        public void Init()
        {
            //string timestamp = DateTime.Now.Ticks.ToString();            
            if (_logUpdateHandler == null)
            {
                _logUpdateHandler = new Logger.LogUpdatedHandler(Logger_LogUpdated);
                Logger.LogUpdated += _logUpdateHandler;
            }

            string gmailUsername;
            string syncProfile;
            string syncContactsFolder;
            string syncAppointmentsFolder;

            GoogleAPITests.LoadSettings(out gmailUsername, out syncProfile, out syncContactsFolder, out syncAppointmentsFolder);

            sync = new Synchronizer();
            sync.SyncAppointments = true;
            sync.SyncContacts = false;
            sync.SyncProfile = syncProfile;
            Assert.IsNotNull(syncAppointmentsFolder);
            AppointmentsSynchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            AppointmentsSynchronizer.TimeMin = DateTime.Now.AddMonths(-1);
            AppointmentsSynchronizer.TimeMax = DateTime.Now.AddMonths(1);

            sync.LoginToGoogle(gmailUsername);
            sync.LoginToOutlook();
        }

        [SetUp]
        public void SetUp()
        {
            DeleteTestAppointments();
        }

        private void DeleteTestAppointments()
        {
            Outlook.MAPIFolder mapiFolder = null;
            Outlook.Items items = null;

            if (string.IsNullOrEmpty(AppointmentsSynchronizer.SyncAppointmentsFolder))
            {
                mapiFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            }
            else
            {
                mapiFolder = Synchronizer.OutlookNameSpace.GetFolderFromID(AppointmentsSynchronizer.SyncAppointmentsFolder);
            }
            Assert.NotNull(mapiFolder);

            try
            {
                items = mapiFolder.Items;
                Assert.NotNull(items);

                object item = items.GetFirst();
                while (item != null)
                {
                    if (item is Outlook.AppointmentItem)
                    {
                        var ola = item as Outlook.AppointmentItem;
                        if (ola.Subject == name)
                        {
                            var s = ola.Subject + " - " + ola.Start;
                            ola.Delete();
                            Logger.Log("Deleted Outlook test appointment: " + s, EventType.Information);
                        }
                        Marshal.ReleaseComObject(ola);
                    }
                    Marshal.ReleaseComObject(item);
                    item = items.GetNext();
                }
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                if (items != null)
                    Marshal.ReleaseComObject(items);
            }

            var query = sync.appointmentsSynchronizer.EventRequest.List(AppointmentsSynchronizer.SyncAppointmentsGoogleFolder);
            Events feed;
            string pageToken = null;
            do
            {
                query.PageToken = pageToken;
                feed = query.Execute();
                foreach (Event e in feed.Items)
                {
                    if (!e.Status.Equals("cancelled") && e.Summary != null && e.Summary == name)
                    {
                        sync.appointmentsSynchronizer.EventRequest.Delete(AppointmentsSynchronizer.SyncAppointmentsGoogleFolder, e.Id).Execute();
                        Logger.Log("Deleted Google test appointment: " + e.Summary + " - " + AppointmentsSynchronizer.GetTime(e), EventType.Information);
                    }
                }
                pageToken = feed.NextPageToken;
            }
            while (pageToken != null);

            sync.appointmentsSynchronizer.LoadAppointments();
            Assert.AreEqual(0, sync.appointmentsSynchronizer.GoogleAppointments.Count);
            Assert.AreEqual(0, sync.appointmentsSynchronizer.OutlookAppointments.Count);
        }

        void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            sync.LogoffOutlook();
            sync.LogoffGoogle();
        }

        [Test]
        public void TestRemoveGoogleDuplicatedAppointments_01()
        {
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var ola1 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola1.Subject = name;
            ola1.Start = DateTime.Now;
            ola1.End = DateTime.Now.AddHours(1);
            ola1.AllDayEvent = false;
            ola1.ReminderSet = false;
            ola1.Save();

            // create new Google test appointments
            var e1 = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(ola1, ref e1);
            var e2 = Factory.NewEvent();
            AppointmentSync.UpdateAppointment(ola1, e2);
            AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(sync.SyncProfile, e2, ola1);
            e2 = sync.appointmentsSynchronizer.SaveGoogleAppointment(e2);

            var gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola1);
            var gid_e1 = AppointmentPropertiesUtils.GetGoogleId(e1);
            var gid_e2 = AppointmentPropertiesUtils.GetGoogleId(e2);
            var oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e1);
            var oid_e2 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e2);
            var oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(ola1);

            // assert appointments ola1 and e1 are pointing to each other
            Assert.AreEqual(gid_ola1, gid_e1);
            Assert.AreEqual(oid_ola1, oid_e1);
            // assert appointment e2 also points to ola1
            Assert.AreEqual(oid_ola1, oid_e2);
            // assert appointment ola1 does not point to e2
            Assert.AreNotEqual(gid_ola1, gid_e2);

            sync.appointmentsSynchronizer.LoadAppointments();

            var f_e1 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e1);
            var f_e2 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e2);
            var f_ola1 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola1);

            Assert.IsNotNull(f_e1);
            Assert.IsNotNull(f_e2);
            Assert.IsNotNull(f_ola1);

            var f_gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola1);
            var f_gid_e1 = AppointmentPropertiesUtils.GetGoogleId(f_e1);
            var f_gid_e2 = AppointmentPropertiesUtils.GetGoogleId(f_e2);
            var f_oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e1);
            var f_oid_e2 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e2);
            var f_oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(f_ola1);

            // assert appointments ola1 and e1 are pointing to each other
            Assert.AreEqual(f_gid_ola1, f_gid_e1);
            Assert.AreEqual(f_oid_ola1, f_oid_e1);
            // assert appointment e2 does not point to ola1
            Assert.AreNotEqual(f_oid_ola1, f_oid_e2);
            // assert appointment ola1 does not point to e2
            Assert.AreNotEqual(f_gid_ola1, f_gid_e2);

            DeleteTestAppointment(f_ola1);
            DeleteTestAppointment(f_e1);
            DeleteTestAppointment(f_e2);
        }

        [Test]
        public void TestRemoveGoogleDuplicatedAppointments_02()
        {
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var ola1 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola1.Subject = name;
            ola1.Start = DateTime.Now;
            ola1.End = DateTime.Now.AddHours(1);
            ola1.AllDayEvent = false;
            ola1.ReminderSet = false;
            ola1.Save();

            // create new Google test appointments
            var e1 = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(ola1, ref e1);
            var e2 = Factory.NewEvent();
            AppointmentSync.UpdateAppointment(ola1, e2);
            AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(sync.SyncProfile, e2, ola1);
            e2 = sync.appointmentsSynchronizer.SaveGoogleAppointment(e2);
            AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola1);

            var gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola1);
            var gid_e1 = AppointmentPropertiesUtils.GetGoogleId(e1);
            var gid_e2 = AppointmentPropertiesUtils.GetGoogleId(e2);
            var oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e1);
            var oid_e2 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e2);
            var oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(ola1);

            // assert appointment e1 points to ola1
            Assert.AreEqual(oid_ola1, oid_e1);
            // assert appointment e2 points to ola1
            Assert.AreEqual(oid_ola1, oid_e2);
            // assert appointment ola1 does not point to e1
            Assert.AreNotEqual(gid_ola1, gid_e1);
            // assert appointment ola1 does not point to e2
            Assert.AreNotEqual(gid_ola1, gid_e2);

            sync.appointmentsSynchronizer.LoadAppointments();

            var f_e1 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e1);
            var f_e2 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e2);
            var f_ola1 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola1);

            Assert.IsNotNull(f_e1);
            Assert.IsNotNull(f_e2);
            Assert.IsNotNull(f_ola1);

            var f_gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola1);
            var f_gid_e1 = AppointmentPropertiesUtils.GetGoogleId(f_e1);
            var f_gid_e2 = AppointmentPropertiesUtils.GetGoogleId(f_e2);
            var f_oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e1);
            var f_oid_e2 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e2);
            var f_oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(f_ola1);

            // assert appointment e1 does not point to ola1
            Assert.AreNotEqual(f_oid_ola1, f_oid_e1);
            // assert appointment ola1 does not point to e1
            Assert.AreNotEqual(f_gid_ola1, f_gid_e1);
            // assert appointment e2 does not point to ola1
            Assert.AreNotEqual(f_oid_ola1, f_oid_e2);
            // assert appointment ola1 does not point to e2
            Assert.AreNotEqual(f_gid_ola1, f_gid_e2);

            DeleteTestAppointment(f_ola1);
            DeleteTestAppointment(f_e1);
            DeleteTestAppointment(f_e2);
        }

        [Test]
        public void TestRemoveOutlookDuplicatedAppointments_01()
        {
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var ola1 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola1.Subject = name;
            ola1.Start = DateTime.Now;
            ola1.End = DateTime.Now.AddHours(1);
            ola1.AllDayEvent = false;
            ola1.ReminderSet = false;
            ola1.Save();

            // create new Google test appointments
            var e1 = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(ola1, ref e1);

            var ola2 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola2.Subject = name;
            ola2.Start = DateTime.Now;
            ola2.End = DateTime.Now.AddHours(1);
            ola2.AllDayEvent = false;
            ola2.ReminderSet = false;
            ola2.Save();
            sync.appointmentsSynchronizer.UpdateAppointment(ola2, ref e1);

            var gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola1);
            var gid_ola2 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola2);
            var gid_e1 = AppointmentPropertiesUtils.GetGoogleId(e1);
            var oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e1);
            var oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(ola1);
            var oid_ola2 = AppointmentPropertiesUtils.GetOutlookId(ola2);

            // assert appointments ola2 and e1 are pointing to each other
            Assert.AreEqual(gid_ola2, gid_e1);
            Assert.AreEqual(oid_ola2, oid_e1);
            // assert appointment ola1 also points to e1
            Assert.AreEqual(gid_ola1, gid_e1);
            // assert appointment e1 does not point to ola1
            Assert.AreNotEqual(oid_e1, oid_ola1);

            sync.appointmentsSynchronizer.LoadAppointments();

            var f_e1 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e1);
            var f_ola1 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola1);
            var f_ola2 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola2);

            Assert.IsNotNull(f_e1);
            Assert.IsNotNull(f_ola1);
            Assert.IsNotNull(f_ola2);

            var f_gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola1);
            var f_gid_ola2 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola2);
            var f_gid_e1 = AppointmentPropertiesUtils.GetGoogleId(f_e1);
            var f_oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e1);
            var f_oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(f_ola1);
            var f_oid_ola2 = AppointmentPropertiesUtils.GetOutlookId(f_ola2);

            // assert appointments ola2 and e1 are pointing to each other
            Assert.AreEqual(f_gid_ola2, f_gid_e1);
            Assert.AreEqual(f_oid_ola2, f_oid_e1);
            // assert appointment ola1 does not point to e1
            Assert.AreNotEqual(f_oid_ola1, f_oid_e1);
            // assert appointment ola1 does not point to e1
            Assert.AreNotEqual(f_gid_ola1, f_gid_e1);

            DeleteTestAppointment(f_ola1);
            DeleteTestAppointment(f_ola2);
            DeleteTestAppointment(f_e1);
        }

        [Test]
        public void TestRemoveOutlookDuplicatedAppointments_02()
        {
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var ola1 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola1.Subject = name;
            ola1.Start = DateTime.Now;
            ola1.End = DateTime.Now.AddHours(1);
            ola1.AllDayEvent = false;
            ola1.ReminderSet = false;
            ola1.Save();

            // create new Google test appointments
            var e1 = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(ola1, ref e1);

            var ola2 = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            ola2.Subject = name;
            ola2.Start = DateTime.Now;
            ola2.End = DateTime.Now.AddHours(1);
            ola2.AllDayEvent = false;
            ola2.ReminderSet = false;
            ola2.Save();
            sync.appointmentsSynchronizer.UpdateAppointment(ola2, ref e1);
            AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(sync.SyncProfile, e1);
            e1 = sync.appointmentsSynchronizer.SaveGoogleAppointment(e1);

            var gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola1);
            var gid_ola2 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, ola2);
            var gid_e1 = AppointmentPropertiesUtils.GetGoogleId(e1);
            var oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, e1);
            var oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(ola1);
            var oid_ola2 = AppointmentPropertiesUtils.GetOutlookId(ola2);

            // assert ola1 points to e1
            Assert.AreEqual(gid_ola1, gid_e1);
            // assert ola2 points to e1
            Assert.AreEqual(gid_ola2, gid_e1);
            // assert appointment e1 does not point to ola1
            Assert.AreNotEqual(oid_e1, oid_ola1);
            // assert appointment e1 does not point to ola2
            Assert.AreNotEqual(oid_e1, oid_ola2);

            sync.appointmentsSynchronizer.LoadAppointments();

            var f_e1 = sync.appointmentsSynchronizer.GetGoogleAppointmentById(gid_e1);
            var f_ola1 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola1);
            var f_ola2 = sync.appointmentsSynchronizer.GetOutlookAppointmentById(oid_ola2);

            Assert.IsNotNull(f_e1);
            Assert.IsNotNull(f_ola1);
            Assert.IsNotNull(f_ola2);

            var f_gid_ola1 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola1);
            var f_gid_ola2 = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync.appointmentsSynchronizer, f_ola2);
            var f_gid_e1 = AppointmentPropertiesUtils.GetGoogleId(f_e1);
            var f_oid_e1 = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, f_e1);
            var f_oid_ola1 = AppointmentPropertiesUtils.GetOutlookId(f_ola1);
            var f_oid_ola2 = AppointmentPropertiesUtils.GetOutlookId(f_ola2);

            // assert ola1 does not point to e1
            Assert.AreNotEqual(f_gid_ola1, f_gid_e1);
            // assert ola2 does not point to e1
            Assert.AreNotEqual(f_gid_ola2, f_gid_e1);
            // assert appointment e1 does not point to ola1
            Assert.AreNotEqual(f_oid_e1, f_oid_ola1);
            // assert appointment e1 does not point to ola2
            Assert.AreNotEqual(f_oid_e1, f_oid_ola2);

            DeleteTestAppointment(f_ola1);
            DeleteTestAppointment(f_ola2);
            DeleteTestAppointment(f_e1);
        }

        [Test]
        public void TestSync_Time()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.End = DateTime.Now.AddHours(1);
            outlookAppointment.AllDayEvent = false;

            outlookAppointment.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var googleAppointment = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(outlookAppointment, ref googleAppointment);

            googleAppointment = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same appointment from google.
            MatchAppointments(sync);
            AppointmentMatch match = FindMatch(outlookAppointment);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);
            Assert.IsNotNull(match.OutlookAppointment);

            Outlook.AppointmentItem recreatedOutlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            sync.appointmentsSynchronizer.UpdateAppointment(ref match.GoogleAppointment, recreatedOutlookAppointment, match.GoogleAppointmentExceptions);
            Assert.IsNotNull(outlookAppointment);
            Assert.IsNotNull(recreatedOutlookAppointment);
            // match recreatedOutlookAppointment with outlookAppointment

            Assert.AreEqual(outlookAppointment.Subject, recreatedOutlookAppointment.Subject);

            Assert.AreEqual(outlookAppointment.Start, recreatedOutlookAppointment.Start);
            Assert.AreEqual(outlookAppointment.End, recreatedOutlookAppointment.End);
            Assert.AreEqual(outlookAppointment.AllDayEvent, recreatedOutlookAppointment.AllDayEvent);
            //ToDo: Check other properties

            DeleteTestAppointments(match);
            recreatedOutlookAppointment.Delete();
        }

        [Test]
        public void TestSync_Day()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.End = DateTime.Now;
            outlookAppointment.AllDayEvent = true;

            outlookAppointment.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var googleAppointment = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(outlookAppointment, ref googleAppointment);

            googleAppointment = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same appointment from google.
            MatchAppointments(sync);
            AppointmentMatch match = FindMatch(outlookAppointment);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);
            Assert.IsNotNull(match.OutlookAppointment);

            Outlook.AppointmentItem recreatedOutlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            sync.appointmentsSynchronizer.UpdateAppointment(ref match.GoogleAppointment, recreatedOutlookAppointment, match.GoogleAppointmentExceptions);
            Assert.IsNotNull(outlookAppointment);
            Assert.IsNotNull(recreatedOutlookAppointment);
            // match recreatedOutlookAppointment with outlookAppointment
            Assert.AreEqual(outlookAppointment.Subject, recreatedOutlookAppointment.Subject);

            Assert.AreEqual(outlookAppointment.Start, recreatedOutlookAppointment.Start);
            Assert.AreEqual(outlookAppointment.End, recreatedOutlookAppointment.End);
            Assert.AreEqual(outlookAppointment.AllDayEvent, recreatedOutlookAppointment.AllDayEvent);
            //ToDo: Check other properties

            DeleteTestAppointments(match);
            recreatedOutlookAppointment.Delete();
        }

        [Test]
        public void TestExtendedProps()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = AppointmentsSynchronizer.CreateOutlookAppointmentItem(AppointmentsSynchronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.AllDayEvent = true;

            outlookAppointment.Save();

            var googleAppointment = Factory.NewEvent();
            sync.appointmentsSynchronizer.UpdateAppointment(outlookAppointment, ref googleAppointment);

            Assert.AreEqual(name, googleAppointment.Summary);

            // read appointment from google
            googleAppointment = null;
            MatchAppointments(sync);
            AppointmentsMatcher.SyncAppointments(sync.appointmentsSynchronizer);

            AppointmentMatch match = FindMatch(outlookAppointment);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);

            // get extended prop
            Assert.AreEqual(AppointmentPropertiesUtils.GetOutlookId(outlookAppointment), AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, match.GoogleAppointment));

            DeleteTestAppointments(match);
        }

        private void DeleteTestAppointments(AppointmentMatch match)
        {
            if (match != null)
            {
                DeleteTestAppointment(match.GoogleAppointment);
                DeleteTestAppointment(match.OutlookAppointment);
            }
        }

        private void DeleteTestAppointment(Outlook.AppointmentItem ola)
        {
            if (ola != null)
            {
                try
                {
                    string name = ola.Subject;
                    ola.Delete();
                    Logger.Log("Deleted Outlook test appointment: " + name, EventType.Information);
                }
                finally
                {
                    Marshal.ReleaseComObject(ola);
                    ola = null;
                }
            }
        }

        private void DeleteTestAppointment(Event e)
        {
            if (e != null && !e.Status.Equals("cancelled"))
            {
                sync.appointmentsSynchronizer.EventRequest.Delete(AppointmentsSynchronizer.SyncAppointmentsGoogleFolder, e.Id);
                Logger.Log("Deleted Google test appointment: " + e.Summary, EventType.Information);
                //Thread.Sleep(2000);
            }
        }

        internal AppointmentMatch FindMatch(Outlook.AppointmentItem ola)
        {
            foreach (AppointmentMatch match in sync.appointmentsSynchronizer.Appointments)
            {
                if (match.OutlookAppointment != null && match.OutlookAppointment.EntryID == ola.EntryID)
                    return match;
            }
            return null;
        }

        private void MatchAppointments(Synchronizer sync)
        {
            //Thread.Sleep(5000); //Wait, until Appointment is really saved and available to retrieve again
            sync.appointmentsSynchronizer.MatchAppointments();
        }

        internal AppointmentMatch FindMatch(Event e)
        {
            if (e != null)
            {
                foreach (AppointmentMatch match in sync.appointmentsSynchronizer.Appointments)
                {
                    if (match.GoogleAppointment != null && match.GoogleAppointment.Id == e.Id)
                        return match;
                }
            }
            return null;
        }
    }
}
