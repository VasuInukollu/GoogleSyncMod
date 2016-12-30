using System;
using NUnit.Framework;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Drawing;
using System.Configuration;
using Google.Contacts;
using Google.GData.Client;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections.ObjectModel;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncContactsTests
    {
        Synchronizer sync;

        static Logger.LogUpdatedHandler _logUpdateHandler = null;

        const int defaultWait = 5000;
        const int defaultWaitTries = 4;

        //Constants for test contact
        const string name = "AN_OUTLOOK_TEST_CONTACT";
        const string email = "email00@outlook.com";
        const string groupName = "A TEST GROUP";
        Group defaultGroup;

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
            sync.SyncContacts = true;
            sync.SyncAppointments = false;
            sync.SyncProfile = syncProfile;
            ContactsSynchronizer.SyncContactsFolder = syncContactsFolder;

            sync.LoginToGoogle(gmailUsername);
            sync.LoginToOutlook();

            //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
            sync.contactsSynchronizer.LoadGoogleGroups();
            defaultGroup = sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup);
        }

        [SetUp]
        public void SetUp()
        {
            // delete previously failed test contacts
            DeleteTestContacts();
            DeleteTestGroups();
            sync.contactsSynchronizer.UseFileAs = true;
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
        public void TestSync_Structured()
        {
            Logger.Log("TestSync_Structured started", EventType.Information);

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.FileAs = name;
            olc1.Email1Address = email;
            olc1.Email2Address = email.Replace("00", "01");
            olc1.Email3Address = email.Replace("00", "02");

            //olc1.HomeAddress = "10 Parades";
            olc1.HomeAddressStreet = "Street";
            olc1.HomeAddressCity = "City";
            olc1.HomeAddressPostalCode = "1234";
            olc1.HomeAddressCountry = "Country";
            olc1.HomeAddressPostOfficeBox = "PO1";
            olc1.HomeAddressState = "State1";

            //olc1.BusinessAddress = "11 Parades"
            olc1.BusinessAddressStreet = "Street2";
            olc1.BusinessAddressCity = "City2";
            olc1.BusinessAddressPostalCode = "5678";
            olc1.BusinessAddressCountry = "Country2";
            olc1.BusinessAddressPostOfficeBox = "PO2";
            olc1.BusinessAddressState = "State2";

            ///olc1.OtherAddress = "12 Parades";
            olc1.OtherAddressStreet = "Street3";
            olc1.OtherAddressCity = "City3";
            olc1.OtherAddressPostalCode = "8012";
            olc1.OtherAddressCountry = "Country3";
            olc1.OtherAddressPostOfficeBox = "PO3";
            olc1.OtherAddressState = "State3";

            #region phones
            //First delete the destination phone numbers
            olc1.PrimaryTelephoneNumber = "123";
            olc1.HomeTelephoneNumber = "456";
            olc1.Home2TelephoneNumber = "4567";
            olc1.BusinessTelephoneNumber = "45678";
            olc1.Business2TelephoneNumber = "456789";
            olc1.MobileTelephoneNumber = "123";
            olc1.BusinessFaxNumber = "1234";
            olc1.HomeFaxNumber = "12345";
            olc1.PagerNumber = "123456";
            //olc1.RadioTelephoneNumber = "1234567";
            olc1.OtherTelephoneNumber = "12345678";
            olc1.CarTelephoneNumber = "123456789";
            olc1.AssistantTelephoneNumber = "987";
            #endregion phones

            #region Name
            olc1.Title = "Title";
            olc1.FirstName = "Firstname";
            olc1.MiddleName = "Middlename";
            olc1.LastName = "Lastname";
            olc1.Suffix = "Suffix";
            //olc1.FullName = name; //The Outlook fullName is automatically set, so don't assign it from Google
            #endregion Name

            olc1.Birthday = new DateTime(1999, 1, 1);

            olc1.NickName = "Nickname";
            olc1.OfficeLocation = "Location";
            olc1.Initials = "IN";
            olc1.Language = "German";

            //olc1.Companies = "Company";
            olc1.CompanyName = "CompanyName";
            olc1.JobTitle = "Position";
            olc1.Department = "Department";

            olc1.IMAddress = "IMs";
            olc1.Anniversary = new DateTime(2000, 1, 1);
            olc1.Children = "Children";
            olc1.Spouse = "Spouse";
            olc1.AssistantName = "Assi";
            olc1.ManagerName = "Chef";
            olc1.WebPage = "http://www.test.de";
            olc1.Body = "<sn>Content & other stuff</sn>\r\n<![CDATA[  \r\n...\r\n&stuff in CDATA < >\r\n  ]]>";
            olc1.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);
            Assert.AreEqual(olc1.Email1Address, olc2.Email1Address);
            Assert.AreEqual(olc1.Email2Address, olc2.Email2Address);
            Assert.AreEqual(olc1.Email3Address, olc2.Email3Address);
            Assert.AreEqual(olc1.PrimaryTelephoneNumber, olc2.PrimaryTelephoneNumber);
            Assert.AreEqual(olc1.HomeTelephoneNumber, olc2.HomeTelephoneNumber);
            Assert.AreEqual(olc1.Home2TelephoneNumber, olc2.Home2TelephoneNumber);
            Assert.AreEqual(olc1.BusinessTelephoneNumber, olc2.BusinessTelephoneNumber);
            Assert.AreEqual(olc1.Business2TelephoneNumber, olc2.Business2TelephoneNumber);
            Assert.AreEqual(olc1.MobileTelephoneNumber, olc2.MobileTelephoneNumber);
            Assert.AreEqual(olc1.BusinessFaxNumber, olc2.BusinessFaxNumber);
            Assert.AreEqual(olc1.HomeFaxNumber, olc2.HomeFaxNumber);
            Assert.AreEqual(olc1.PagerNumber, olc2.PagerNumber);
            //Assert.AreEqual(olc1.RadioTelephoneNumber, olc2.RadioTelephoneNumber);
            Assert.AreEqual(olc1.OtherTelephoneNumber, olc2.OtherTelephoneNumber);
            Assert.AreEqual(olc1.CarTelephoneNumber, olc2.CarTelephoneNumber);
            Assert.AreEqual(olc1.AssistantTelephoneNumber, olc2.AssistantTelephoneNumber);

            Assert.AreEqual(olc1.HomeAddressStreet, olc2.HomeAddressStreet);
            Assert.AreEqual(olc1.HomeAddressCity, olc2.HomeAddressCity);
            Assert.AreEqual(olc1.HomeAddressCountry, olc2.HomeAddressCountry);
            Assert.AreEqual(olc1.HomeAddressPostalCode, olc2.HomeAddressPostalCode);
            Assert.AreEqual(olc1.HomeAddressPostOfficeBox, olc2.HomeAddressPostOfficeBox);
            Assert.AreEqual(olc1.HomeAddressState, olc2.HomeAddressState);

            Assert.AreEqual(olc1.BusinessAddressStreet, olc2.BusinessAddressStreet);
            Assert.AreEqual(olc1.BusinessAddressCity, olc2.BusinessAddressCity);
            Assert.AreEqual(olc1.BusinessAddressCountry, olc2.BusinessAddressCountry);
            Assert.AreEqual(olc1.BusinessAddressPostalCode, olc2.BusinessAddressPostalCode);
            Assert.AreEqual(olc1.BusinessAddressPostOfficeBox, olc2.BusinessAddressPostOfficeBox);
            Assert.AreEqual(olc1.BusinessAddressState, olc2.BusinessAddressState);

            Assert.AreEqual(olc1.OtherAddressStreet, olc2.OtherAddressStreet);
            Assert.AreEqual(olc1.OtherAddressCity, olc2.OtherAddressCity);
            Assert.AreEqual(olc1.OtherAddressCountry, olc2.OtherAddressCountry);
            Assert.AreEqual(olc1.OtherAddressPostalCode, olc2.OtherAddressPostalCode);
            Assert.AreEqual(olc1.OtherAddressPostOfficeBox, olc2.OtherAddressPostOfficeBox);
            Assert.AreEqual(olc1.OtherAddressState, olc2.OtherAddressState);

            Assert.AreEqual(olc1.FullName, olc2.FullName);
            Assert.AreEqual(olc1.MiddleName, olc2.MiddleName);
            Assert.AreEqual(olc1.LastName, olc2.LastName);
            Assert.AreEqual(olc1.FirstName, olc2.FirstName);
            Assert.AreEqual(olc1.Title, olc2.Title);
            Assert.AreEqual(olc1.Suffix, olc2.Suffix);

            Assert.AreEqual(olc1.Birthday, olc2.Birthday);

            Assert.AreEqual(olc1.NickName, olc2.NickName);
            Assert.AreEqual(olc1.OfficeLocation, olc2.OfficeLocation);
            Assert.AreEqual(olc1.Initials, olc2.Initials);
            Assert.AreEqual(olc1.Language, olc2.Language);

            Assert.AreEqual(olc1.IMAddress, olc2.IMAddress);
            Assert.AreEqual(olc1.Anniversary, olc2.Anniversary);
            Assert.AreEqual(olc1.Children, olc2.Children);
            Assert.AreEqual(olc1.Spouse, olc2.Spouse);
            Assert.AreEqual(olc1.ManagerName, olc2.ManagerName);
            Assert.AreEqual(olc1.AssistantName, olc2.AssistantName);

            Assert.AreEqual(olc1.WebPage, olc2.WebPage);
            Assert.AreEqual(olc1.Body, olc2.Body);

            //Assert.AreEqual(olc1.Companies, olc2.Companies); 
            Assert.AreEqual(olc1.CompanyName, olc2.CompanyName);
            Assert.AreEqual(olc1.JobTitle, olc2.JobTitle);
            Assert.AreEqual(olc1.Department, olc2.Department);

            DeleteAppointmentsForTestContacts();

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_Structured finished", EventType.Information);
        }

        [Test]
        public void TestSync_Unstructured()
        {
            Logger.Log("TestSync_Unstructured started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.FileAs = name;
            olc1.HomeAddress = "10 Parades";
            olc1.BusinessAddress = "11 Parades";
            olc1.OtherAddress = "12 Parades";
            olc1.IMAddress = "  "; //Test empty IMAddress
            olc1.Email2Address = "  "; //Test empty Email Address
            olc1.FullName = name;
            olc1.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);

            Assert.AreEqual(olc1.HomeAddress, olc2.HomeAddress);
            Assert.AreEqual(olc1.BusinessAddress, olc2.BusinessAddress);
            Assert.AreEqual(olc1.OtherAddress, olc2.OtherAddress);

            Assert.AreEqual(olc1.FullName, olc2.FullName);

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_Unstructured finished", EventType.Information);
        }

        [Test]
        public void TestSync_CompanyOnly()
        {
            Logger.Log("TestSync_CompanyOnly started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.CompanyName = name;
            olc1.BusinessAddress = "11 Parades";
            olc1.Save();

            Assert.IsNull(olc1.FullName);
            Assert.IsNull(olc1.Email1Address);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);

            Assert.AreEqual(olc1.CompanyName, olc2.CompanyName);

            Assert.AreEqual(olc1.BusinessAddress, olc2.BusinessAddress);

            Assert.IsNull(olc2.FullName);
            Assert.IsNull(olc1.Email1Address);

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_CompanyOnly finished", EventType.Information);
        }

        [Test]
        public void TestSync_EmailOnly()
        {
            Logger.Log("TestSync_EmailOnly started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.FileAs = email;
            olc1.Email1Address = email;
            olc1.Save();

            Assert.IsNull(olc1.FullName);
            Assert.IsNull(olc1.CompanyName);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(email, email);
            //ContactsMatcher.SyncContact(match, sync);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);

            Assert.AreEqual(olc1.Email1Address, olc2.Email1Address);

            Assert.IsNull(olc2.FullName);
            Assert.IsNull(olc1.CompanyName);

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_EmailOnly finished", EventType.Information);
        }

        [Test]
        public void TestSync_UseFileAs()
        {
            Logger.Log("TestSync_UseFileAs started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.contactsSynchronizer.UseFileAs = true;

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.FullName = name;
            olc1.FileAs = "SaveAs";
            olc1.Save();

            Assert.AreNotEqual(olc1.FullName, olc1.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty("SaveAs", string.Empty);
            //ContactsMatcher.SyncContact(match, sync);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            Assert.IsNotNull(match.GoogleContact);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(olc2.FileAs, match.GoogleContact.Title);
            Assert.AreEqual(olc2.FileAs, match.GoogleContact.Name.FullName);
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);

            olc2.FileAs = name;
            Assert.AreNotEqual(olc1.FileAs, olc2.FileAs);
            Assert.AreNotEqual(olc2.FileAs, match.GoogleContact.Title);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);
            Assert.AreEqual(match.GoogleContact.Name.FamilyName, olc2.FileAs);

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_UseFileAs finished", EventType.Information);
        }

        [Test]
        public void TestSync_UseFullName()
        {
            Logger.Log("TestSync_UseFullName started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.contactsSynchronizer.UseFileAs = false;

            // create new contact to sync
            var olc1 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc1.FullName = name;
            olc1.FileAs = "SaveAs";
            olc1.Save();

            Assert.AreNotEqual(olc1.FullName, olc1.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact gc = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc1, gc);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc1, sync.contactsSynchronizer), gc);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match);

            var olc2 = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);

            // match olc2 with olc1
            Assert.AreEqual(match.GoogleContact.Name.FullName, OutlookContactInfo.GetTitleFirstLastAndSuffix(olc2).Trim().Replace("  ", " "));
            Assert.AreNotEqual(olc1.FileAs, match.GoogleContact.Title);
            Assert.AreNotEqual(olc1.FileAs, match.GoogleContact.Name.FullName);
            Assert.AreNotEqual(olc1.FileAs, olc2.FileAs);

            olc2.FileAs = "SaveAs";
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);
            ContactSync.UpdateContact(match.GoogleContact, olc2, sync.contactsSynchronizer.UseFileAs);
            Assert.AreEqual(olc1.FileAs, olc2.FileAs);

            olc1.Delete();
            Marshal.ReleaseComObject(olc1);
            olc1 = null;

            olc2.Delete();
            Marshal.ReleaseComObject(olc2);
            olc2 = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSync_UseFullName finished", EventType.Information);
        }

        [Test]
        public void TestExtendedProps()
        {
            Logger.Log("TestExtendedProps started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.contactsSynchronizer.UseFileAs = true;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            var gc = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, gc);
            var m = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), gc);

            sync.contactsSynchronizer.SaveGoogleContact(m);

            Assert.AreEqual(name, m.GoogleContact.Title);

            // read contact from google
            Assert.IsTrue(EnsureGoogleContactSaved(m.GoogleContact));
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncContacts(sync.contactsSynchronizer);

            m = sync.contactsSynchronizer.ContactByProperty(name, email);

            Assert.IsNotNull(m);
            Assert.IsNotNull(m.GoogleContact);

            // get extended prop
            var ooid = ContactPropertiesUtils.GetOutlookId(olc);
            var goid = ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, m.GoogleContact);
            Assert.AreEqual(ooid, goid);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(m.GoogleContact);

            Logger.Log("TestExtendedProps finished", EventType.Information);
        }

        [Test]
        public void TestSyncDeletedOulook()
        {
            Logger.Log("TestSyncDeletedOulook started", EventType.Information);

            sync.contactsSynchronizer.LoadContacts();
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.AreEqual(0, sync.contactsSynchronizer.OutlookContacts.Count);

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contacts
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            // delete outlook contact
            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            // sync
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncContacts(sync.contactsSynchronizer);
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleContact);
            Assert.IsNull(match.OutlookContact);

            // delete
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactDeleted(match.GoogleContact));

            // sync
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncContacts(sync.contactsSynchronizer);

            // check if google contact still exists
            match = sync.contactsSynchronizer.ContactByProperty(name, email);

            Assert.IsNull(match);

            Logger.Log("TestSyncDeletedOulook finished", EventType.Information);
        }

        [Test]
        public void TestSyncDeletedGoogle()
        {
            Logger.Log("TestSyncDeletedGoogle started", EventType.Information);

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contacts
            sync.contactsSynchronizer.SaveContact(match);

            // delete google contact
            sync.contactsSynchronizer.ContactsRequest.Delete(match.GoogleContact);

            // sync
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // delete
            sync.contactsSynchronizer.SaveContact(match);

            // sync
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncContacts(sync.contactsSynchronizer);
            match = sync.contactsSynchronizer.ContactByProperty(name, email);

            // check if outlook contact still exists
            Assert.IsNull(match);

            //deleted in test case olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            Logger.Log("TestSyncDeletedGoogle finished", EventType.Information);
        }

        [Test]
        public void TestGooglePhoto()
        {
            Logger.Log("TestGooglePhoto started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            Image pic = Utilities.CropImageGoogleFormat(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));
            bool saved = Utilities.SaveGooglePhoto(sync.contactsSynchronizer, match.GoogleContact, pic);
            Assert.IsTrue(saved);
            Assert.IsTrue(EnsureGoogleContactHasPhoto(match.GoogleContact));

            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            Image image = Utilities.GetGooglePhoto(sync.contactsSynchronizer, match.GoogleContact);
            Assert.IsNotNull(image);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestGooglePhoto finished", EventType.Information);
        }

        [Test]
        public void TestOutlookPhoto()
        {
            Logger.Log("TestOutlookPhoto started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            bool saved = Utilities.SetOutlookPhoto(olc, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            Assert.IsTrue(saved);

            olc.Save();

            Image image = Utilities.GetOutlookPhoto(olc);
            Assert.IsNotNull(image);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            Logger.Log("TestOutlookPhoto finished", EventType.Information);
        }

        [Test]
        public void TestSyncPhoto()
        {
            Logger.Log("TestSyncPhoto started", EventType.Information);

            sync.contactsSynchronizer.LoadContacts();
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.AreEqual(0, sync.contactsSynchronizer.OutlookContacts.Count);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            Utilities.SetOutlookPhoto(olc, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            olc.Save();

            // outlook contact should now have a photo
            Assert.IsNotNull(Utilities.GetOutlookPhoto(olc));

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // google contact should now have a photo
            Assert.IsNotNull(Utilities.GetGooglePhoto(sync.contactsSynchronizer, match.GoogleContact));

            // delete outlook contact
            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            // recreate outlook contact
            olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);

            // outlook contact should now have no photo
            Assert.IsNull(Utilities.GetOutlookPhoto(olc));

            sync.contactsSynchronizer.UpdateContact(match.GoogleContact, olc);
            match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), match.GoogleContact);
            //match.OutlookContact.Save();            

            //save contact to outlook
            sync.contactsSynchronizer.SaveContact(match);

            // outlook contact should now have a photo
            Assert.IsNotNull(Utilities.GetOutlookPhoto(olc));

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSyncPhoto finished", EventType.Information);
        }

        [Test]
        public void TestSyncGroups()
        {
            Logger.Log("TestSyncGroups started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Categories = groupName;
            olc.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, olc.Categories);

            //Sync Groups first
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncGroups(sync.contactsSynchronizer);

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //sync and save contact to google.
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // google contact should now have the same group
            var googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Assert.AreEqual(2, googleGroups.Count);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(groupName), googleGroups);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);

            // delete outlook contact
            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            sync.contactsSynchronizer.UpdateContact(match.GoogleContact, olc);
            match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), match.GoogleContact);
            olc.Save();

            sync.SyncOption = SyncOption.MergeGoogleWins;

            //sync and save contact to outlook
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);
            sync.contactsSynchronizer.SaveContact(match);

            //load the same contact from outlook
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            Assert.AreEqual(groupName, olc.Categories);

            // delete test group
            Group group = sync.contactsSynchronizer.GetGoogleGroupByName(groupName);
            if (group != null)
                sync.contactsSynchronizer.ContactsRequest.Delete(group);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSyncGroups finished", EventType.Information);
        }

        [Test]
        public void TestSyncDeletedGoogleGroup()
        {
            Logger.Log("TestSyncDeletedGoogleGroup started", EventType.Information);

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Categories = groupName;
            olc.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, olc.Categories);

            //Sync Groups first
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncGroups(sync.contactsSynchronizer);
            Assert.IsTrue(EnsureGoogleGroupSaved(groupName));

            //Create now Google Contact and assing new Group
            var googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            var match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contact to google.            
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));
            Assert.AreEqual(2, match.GoogleContact.GroupMembership.Count);

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match.GoogleContact);
            Assert.IsNotNull(match.OutlookContact);
            Assert.AreEqual(2, match.GoogleContact.GroupMembership.Count);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // google contact should now have the same group
            var googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Assert.AreEqual(2, googleGroups.Count);

            var group = sync.contactsSynchronizer.GetGoogleGroupByName(groupName);
            Assert.Contains(group, googleGroups);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);

            // delete group from google contact
            Utilities.RemoveGoogleGroup(match.GoogleContact, group);

            googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);

            //save contact to google.
            sync.contactsSynchronizer.SaveGoogleContact(match.GoogleContact);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));
            Assert.IsTrue(EnsureGoogleContactHasGroups(match.GoogleContact, 1));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;

            //Sync Groups first
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncGroups(sync.contactsSynchronizer);

            //sync and save contact to outlook.
            var etag = match.GoogleContact.ETag;
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            sync.contactsSynchronizer.UpdateContact(match.GoogleContact, olc);
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactUpdated(match.GoogleContact, etag));
            Assert.AreEqual(1, match.GoogleContact.GroupMembership.Count);

            // google and outlook should now have no category except for the System Group: My Contacts
            googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.AreEqual(null, olc.Categories);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);

            // delete test group
            if (group != null)
                sync.contactsSynchronizer.ContactsRequest.Delete(group);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSyncDeletedGoogleGroup finished", EventType.Information);
        }

        [Test]
        public void TestSyncDeletedOutlookGroup()
        {
            Logger.Log("TestSyncDeletedOutlookGroup started", EventType.Information);

            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Categories = groupName;
            olc.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, olc.Categories);

            //Now sync Groups
            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncGroups(sync.contactsSynchronizer);

            Contact googleContact = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), googleContact);

            //save contact to google.
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // google contact should now have the same group
            Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Group group = sync.contactsSynchronizer.GetGoogleGroupByName(groupName);
            Assert.AreEqual(2, googleGroups.Count);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);
            Assert.Contains(group, googleGroups);

            // delete group from outlook
            Utilities.RemoveOutlookGroup(olc, groupName);

            //save contact to google.
            sync.contactsSynchronizer.SaveContact(match);

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            sync.contactsSynchronizer.UpdateContact(olc, match.GoogleContact);

            // google and outlook should now have no category
            googleGroups = Utilities.GetGoogleGroups(sync.contactsSynchronizer, match.GoogleContact);
            Assert.AreEqual(null, olc.Categories);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.Contains(sync.contactsSynchronizer.GetGoogleGroupByName(ContactsSynchronizer.myContactsGroup), googleGroups);

            // delete test group
            if (group != null)
                sync.contactsSynchronizer.ContactsRequest.Delete(group);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestSyncDeletedOutlookGroup finished", EventType.Information);
        }

        [Test]
        public void TestResetMatches()
        {
            Logger.Log("TestResetMatches started", EventType.Information);

            sync.SyncOption = SyncOption.MergeOutlookWins;
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);

            // create new contact to sync
            var olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            var gc = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, gc);
            var match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), gc);

            //save contact to google.
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);
            
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match.GoogleContact);
            Assert.IsNotNull(match.OutlookContact);
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);
            Assert.IsFalse(sync.SyncDelete);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);
            
            // delete outlook contact
            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            //load the same contact from google
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);

            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);
            Assert.IsFalse(sync.SyncDelete);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // reset matches
            var etag = match.GoogleContact.ETag;
            sync.contactsSynchronizer.ResetMatch(match.GoogleContact);
            Assert.IsTrue(EnsureGoogleContactUpdated(match.GoogleContact, etag));
            //Not, because NULL: sync.ResetMatch(match.OutlookContact.GetOriginalItemFromOutlook(sync));

            // load same contact match
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);
            Assert.IsFalse(sync.SyncDelete);

            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // google contact should still be present and OutlookContact should be filled
            Assert.IsNotNull(match.GoogleContact);
            Assert.IsNotNull(match.OutlookContact);

            DeleteTestContacts(match);

             // create new contact to sync
            olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
            olc.FullName = name;
            olc.FileAs = name;
            olc.Email1Address = email;
            olc.Email2Address = email.Replace("00", "01");
            olc.Email3Address = email.Replace("00", "02");
            olc.HomeAddress = "10 Parades";
            olc.PrimaryTelephoneNumber = "123";
            olc.Save();

            // same test for delete google contact...
            gc = new Contact();
            sync.contactsSynchronizer.UpdateContact(olc, gc);
            match = new ContactMatch(new OutlookContactInfo(olc, sync.contactsSynchronizer), gc);

            //save contact to google.
            sync.contactsSynchronizer.SaveContact(match);
            Assert.IsTrue(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);

            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);
            Assert.IsFalse(sync.SyncDelete);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);

            // delete google contact           
            sync.contactsSynchronizer.ContactsRequest.Delete(match.GoogleContact);
            Assert.IsTrue(EnsureGoogleContactDeleted(match.GoogleContact));            

            //load the same contact from google.
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);

            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNull(match.GoogleContact);

            // reset matches
            //Not, because null: sync.ResetMatch(match.GoogleContact);
            sync.contactsSynchronizer.ResetMatch(match.OutlookContact.GetOriginalItemFromOutlook());

            // load same contact match
            sync.contactsSynchronizer.MatchContacts();
            Assert.IsNotNull(sync.contactsSynchronizer.GoogleContacts);
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.OutlookContacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.OutlookContacts.Count);
            Assert.IsNotNull(sync.contactsSynchronizer.Contacts);
            Assert.AreEqual(1, sync.contactsSynchronizer.Contacts.Count);
            match = sync.contactsSynchronizer.ContactByProperty(name, email);
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNull(match.GoogleContact);
            Assert.AreEqual(SyncOption.MergeOutlookWins, sync.SyncOption);
            Assert.IsFalse(sync.SyncDelete);
            ContactsMatcher.SyncContact(match, sync.contactsSynchronizer);

            // Outlook contact should still be present and GoogleContact should be filled
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);

            olc.Delete();
            Marshal.ReleaseComObject(olc);
            olc = null;

            DeleteTestContact(match.GoogleContact);

            Logger.Log("TestResetMatches finished", EventType.Information);
        }

        private void DeleteTestContacts(ContactMatch match)
        {
            if (match != null)
            {
                DeleteTestContact(match.GoogleContact);
                DeleteTestContact(match.OutlookContact);
            }
        }

        private void DeleteTestContact(Outlook.ContactItem c)
        {
            if (c != null)
            {
                try
                {
                    string name = c.FileAs;
                    c.Delete();
                    Logger.Log("Deleted Outlook test contact: " + name, EventType.Information);
                }
                finally
                {
                    Marshal.ReleaseComObject(c);
                    c = null;
                }
            }
        }

        private void DeleteTestContact(OutlookContactInfo outlookContact)
        {
            if (outlookContact != null)
                DeleteTestContact(outlookContact.GetOriginalItemFromOutlook());
        }

        private void DeleteTestContact(Contact c1)
        {
            if (c1 != null && !c1.Deleted && c1.AtomEntry.EditUri != null)
            {
                try
                {
                    sync.contactsSynchronizer.ContactsRequest.Delete(c1);
                    Logger.Log("Deleted Google test contact: " + c1.Title, EventType.Information);
                }
                catch (GDataVersionConflictException)
                {
                    try
                    {
                        var c2 = sync.contactsSynchronizer.ContactsRequest.Retrieve<Contact>(new Uri(c1.Self));
                        if (c2 != null && !c2.Deleted)
                        {
                            sync.contactsSynchronizer.ContactsRequest.Delete(c2);
                            Logger.Log("Deleted Google test contact: " + c2.Title, EventType.Information);
                        }
                    }
                    catch (Exception e1)
                    {
                        Logger.Log(e1, EventType.Information);
                    }
                }
                catch (Exception e2)
                {
                    Logger.Log(e2, EventType.Information);
                }
            }
        }

        private void DeleteTestGroup(Group g)
        {
            if (g != null && !g.Deleted)
            {
                try
                {
                    sync.contactsSynchronizer.ContactsRequest.Delete(g);
                    Logger.Log("Deleted Google test group: " + g.Title, EventType.Information);
                }
                catch (Exception e1)
                {
                    Logger.Log(e1, EventType.Information);
                }
            }
        }

        [Ignore("TestMassSyncToGoogle")]
        public void TestMassSyncToGoogle()
        {
            // NEED TO DELETE CONTACTS MANUALY

            int c = 300;
            string[] names = new string[c];
            string[] emails = new string[c];
            Outlook.ContactItem outlookContact;
            ContactMatch match;

            for (int i = 0; i < c; i++)
            {
                names[i] = "TEST name" + i;
                emails[i] = "email" + i + "@domain.com";
            }

            // count existing google contacts
            int excount = sync.contactsSynchronizer.GoogleContacts.Count;

            DateTime start = DateTime.Now;
            Console.WriteLine("Started mass sync to google of " + c + " contacts");

            for (int i = 0; i < c; i++)
            {
                outlookContact = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);
                outlookContact.FullName = names[i];
                outlookContact.FileAs = names[i];
                outlookContact.Email1Address = emails[i];
                outlookContact.Save();

                Contact googleContact = new Contact();
                ContactSync.UpdateContact(outlookContact, googleContact, sync.contactsSynchronizer.UseFileAs);
                match = new ContactMatch(new OutlookContactInfo(outlookContact, sync.contactsSynchronizer), googleContact);

                //save contact to google.
                sync.contactsSynchronizer.SaveGoogleContact(match);
            }

            sync.contactsSynchronizer.MatchContacts();
            ContactsMatcher.SyncContacts(sync.contactsSynchronizer);

            // all contacts should be synced
            Assert.AreEqual(c, sync.contactsSynchronizer.Contacts.Count - excount);

            DateTime end = DateTime.Now;
            TimeSpan time = end - start;
            Console.WriteLine("Synced " + c + " contacts to google in " + time.TotalSeconds + " s ("
                + ((float)time.TotalSeconds / c) + " s per contact)");

            // received: Synced 50 contacts to google in 30.137 s (0.60274 s per contact)
        }

        [Ignore("TestCreatingGoogeAccountThatFailed1")]
        public void TestCreatingGoogeAccountThatFailed1()
        {
            Outlook.ContactItem outlookContact = sync.contactsSynchronizer.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(olc1, gc);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            ContactSync.SetAddresses(outlookContact, googleContact);

            ContactSync.SetCompanies(outlookContact, googleContact);

            ContactSync.SetIMs(outlookContact, googleContact);

            googleContact.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = sync.contactsSynchronizer.ContactsRequest.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync.contactsSynchronizer, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
        }

        //[Test]
        [Ignore("TestCreatingGoogeAccountThatFailed2")]
        public void TestCreatingGoogeAccountThatFailed2()
        {
            Outlook.ContactItem outlookContact = sync.contactsSynchronizer.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(olc1, gc);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            //SetPhoneNumbers(olc1, gc);

            //SetAddresses(olc1, gc);

            //SetCompanies(olc1, gc);

            //SetIMs(olc1, gc);

            googleContact.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = sync.contactsSynchronizer.ContactsRequest.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync.contactsSynchronizer, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
        }

        //[Test]
        [Ignore("TestCreatingGoogeAccountThatFailed3")]
        public void TestCreatingGoogeAccountThatFailed3()
        {
            Outlook.ContactItem outlookContact = sync.contactsSynchronizer.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(olc1, gc);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            //SetEmails(olc1, gc);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(olc1, gc);

            //SetCompanies(olc1, gc);

            //SetIMs(olc1, gc);

            //gc.Content.Content = olc1.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = sync.contactsSynchronizer.ContactsRequest.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync.contactsSynchronizer, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
        }

        //[Test]
        [Ignore("TestUpdatingGoogeAccountThatFailed")]
        public void TestUpdatingGoogeAccountThatFailed()
        {
            Outlook.ContactItem outlookContact = sync.contactsSynchronizer.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleContact);

            Contact googleContact = match.GoogleContact;

            ContactSync.UpdateContact(outlookContact, googleContact, sync.contactsSynchronizer.UseFileAs);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(olc1, gc);

            //SetCompanies(olc1, gc);

            //SetIMs(olc1, gc);

            //gc.Content.Content = olc1.Body;

            Contact updatedEntry = sync.contactsSynchronizer.ContactsRequest.Update(googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync.contactsSynchronizer, outlookContact, updatedEntry);
            match.GoogleContact = updatedEntry;
            outlookContact.Save();
        }

        internal ContactMatch FindMatch(Outlook.ContactItem outlookContact)
        {
            foreach (ContactMatch match in sync.contactsSynchronizer.Contacts)
            {
                if (match.OutlookContact != null && match.OutlookContact.EntryID == outlookContact.EntryID)
                    return match;
            }
            return null;
        }

        private void DeleteGoogleTestContacts()
        {
            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;
            query.Group = defaultGroup.Id;

            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);
            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (IsTestContact(a))
                    {
                        DeleteTestContact(a);
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
        }

        private void DeleteOutlookTestContacts()
        {
            Outlook.MAPIFolder mapiFolder = null;
            Outlook.Items items = null;

            if (string.IsNullOrEmpty(ContactsSynchronizer.SyncContactsFolder))
            {
                mapiFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            }
            else
            {
                mapiFolder = Synchronizer.OutlookNameSpace.GetFolderFromID(ContactsSynchronizer.SyncContactsFolder);
            }
            Assert.NotNull(mapiFolder);

            try
            {
                items = mapiFolder.Items;
                Assert.NotNull(items);

                object item = items.GetFirst();
                while (item != null)
                {
                    if (item is Outlook.ContactItem)
                    {
                        var olc = item as Outlook.ContactItem;
                        if (IsTestContact(olc))
                        {
                            var s = olc.FullName;
                            olc.Delete();
                            Logger.Log("Deleted Outlook test contact: " + s, EventType.Information);
                        }
                        Marshal.ReleaseComObject(olc);
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
        }

        private void DeleteTestContacts()
        {
            DeleteOutlookTestContacts();

            for (var i = 1; i < defaultWaitTries; i++)
            {
                DeleteGoogleTestContacts();
                if (EnsureAllGoogleContactsDeleted())
                    break;
            }

            sync.contactsSynchronizer.LoadContacts();
            Assert.AreEqual(0, sync.contactsSynchronizer.GoogleContacts.Count);
            Assert.AreEqual(0, sync.contactsSynchronizer.OutlookContacts.Count);
        }

        private bool EnsureAllGoogleContactsDeleted()
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (!FindAnyTestContacts())
                {
                    sync.contactsSynchronizer.LoadContacts();
                    if (sync.contactsSynchronizer.GoogleContacts.Count == 0)
                        return true;
                }
                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureAllGoogleContactsDeleted: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactHasPhoto(Contact c)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;
                query.Group = defaultGroup.Id;

                var id = c.ContactEntry.Id;
                var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);

                while (feed != null)
                {
                    foreach (var a in feed.Entries)
                    {
                        if (id.Equals(a.ContactEntry.Id))
                        {
                            if (Utilities.HasPhoto(a))
                                return true;
                        }
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleContactHasPhoto: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }

            return false;
        }

        private bool EnsureGoogleContactHasGroups(Contact c, int groups)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;
                query.Group = defaultGroup.Id;

                var id = c.ContactEntry.Id;
                var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);

                while (feed != null)
                {
                    foreach (var a in feed.Entries)
                    {
                        if (id.Equals(a.ContactEntry.Id))
                        {
                            if (a.GroupMembership.Count == groups)
                                return true;
                        }
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleContactHasGroups: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }

            return false;
        }

        private bool EnsureGoogleGroupSaved(string gn)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrieveGroup(gn))
                    return true;
                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleGroupSaved: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactSaved(Contact c)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrieveContact(c))
                    return true;
                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleContactSaved: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactDeleted(Contact c)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (!RetrieveContact(c))
                    return true;
                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleContactDeleted: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactUpdated(Contact c, string etag)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrieveContactIfUpdated(c, etag))
                    return true;
                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Logger.Log("EnsureGoogleContactUpdated: sleeping for " + t + "ms", EventType.Information);
                Thread.Sleep(t);
            }
            return false;
        }

        private bool RetrieveContactIfUpdated(Contact c, string etag)
        {
            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;
            query.Group = defaultGroup.Id;

            var id = c.ContactEntry.Id;
            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);

            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (id.Equals(a.ContactEntry.Id) && (a.ContactEntry.Etag != etag))
                        return true;
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
            return false;
        }

        private bool RetrieveContact(Contact c)
        {
            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;
            query.Group = defaultGroup.Id;

            var id = c.ContactEntry.Id;
            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);

            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (id.Equals(a.ContactEntry.Id))
                        return true;
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
            return false;
        }

        private bool RetrieveGroup(string gn)
        {
            var query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;

            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Group>(query);
            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (a.Title == gn)
                    {
                        return true;
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
            return false;
        }

        private bool FindAnyTestContacts()
        {
            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;
            query.Group = defaultGroup.Id;

            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Contact>(query);
            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (IsTestContact(a))
                    {
                        return true;
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
            return false;
        }

        private void DeleteTestGroups()
        {
            var query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;

            var feed = sync.contactsSynchronizer.ContactsRequest.Get<Group>(query);
            while (feed != null)
            {
                foreach (var a in feed.Entries)
                {
                    if (IsTestGroup(a))
                    {
                        DeleteTestGroup(a);
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = sync.contactsSynchronizer.ContactsRequest.Get(feed, FeedRequestType.Next);
            }
        }

        void DeleteAppointmentsForTestContacts()
        {
            //Also delete the birthday/anniversary entries in Outlook calendar
            Logger.Log("Deleting Outlook calendar TEST entries (birthday, anniversary) ...", EventType.Information);

            try
            {
                var outlookNamespace = Synchronizer.OutlookApplication.GetNamespace("mapi");
                var calendarFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var outlookCalendarItems = calendarFolder.Items;
                for (int i = outlookCalendarItems.Count; i > 0; i--)
                {
                    var item = outlookCalendarItems[i] as Outlook.AppointmentItem;
                    if (item.Subject.Contains(name))
                    {
                        string subject = item.Subject;
                        item.Delete();
                        Marshal.ReleaseComObject(item);
                        item = null;
                        Logger.Log("Deleted Outlook calendar TEST entry: " + subject, EventType.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Could not delete Outlook calendar TEST entries: " + ex.Message, EventType.Information);
            }
        }

        private bool IsTestContact(Outlook.ContactItem olc)
        {
            if (olc.Email1Address == email)
                return true;

            if (olc.FileAs == name)
                return true;

            if (olc.FileAs == "SaveAs")
                return true;

            return false;
        }

        private bool IsTestGroup(Group g)
        {
            if (g.Title == groupName)
            {
                return true;
            }
            return false;
        }

        private bool IsTestContact(Contact gc)
        {
            if (gc == null)
                return false;

            if (gc.PrimaryEmail != null && gc.PrimaryEmail.Address == email)
                return true;

            if (gc.Title == name)
                return true;

            if (gc.Name != null)
            {
                if (gc.Name.FullName == name)
                    return true;
                if (gc.Name.FamilyName == name)
                    return true;
            }

            if (gc.Organizations != null && gc.Organizations.Count > 0)
            {
                if (gc.Organizations[0].Name == name)
                    return true;
            }

            return false;
        }
    }
}
