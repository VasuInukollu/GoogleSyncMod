using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using Google.Contacts;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    internal static class ContactPropertiesUtils
    {
        public static string GetOutlookId(Outlook.ContactItem outlookContact)
        {
            return outlookContact.EntryID;
        }

        public static string GetGoogleId(Contact googleContact)
        {
            string id = googleContact.Id.ToString();
            if (id == null)
                throw new Exception();
            return id;
        }

        public static void SetGoogleOutlookContactId(string syncProfile, Contact googleContact, Outlook.ContactItem outlookContact)
        {
            if (outlookContact.EntryID == null)
                throw new Exception("Must save outlook contact before getting id");

            SetGoogleOutlookContactId(syncProfile, googleContact, GetOutlookId(outlookContact));
        }

        public static void SetGoogleOutlookContactId(string syncProfile, Contact googleContact, string outlookContactId)
        {
            // check if exists
            bool found = false;
            foreach (var p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                {
                    p.Value = outlookContactId;
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                var prop = new ExtendedProperty(outlookContactId, "gos:oid:" + syncProfile + "");
                prop.Value = outlookContactId;
                googleContact.ExtendedProperties.Add(prop);
            }
        }

        public static string GetGoogleOutlookContactId(string syncProfile, Contact googleContact)
        {
            // get extended prop
            foreach (var p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                    return p.Value;
            }
            return null;
        }

        public static void ResetGoogleOutlookContactId(string syncProfile, Contact googleContact)
        {
            // get extended prop
            foreach (var p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                {
                    // remove 
                    googleContact.ExtendedProperties.Remove(p);
                    return;
                }
            }
        }

        /// <summary>
        /// Sets the syncId of the Outlook contact and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="outlookContact"></param>
        /// <param name="googleContact"></param>
        public static void SetOutlookGoogleContactId(ContactsSynchronizer sync, Outlook.ContactItem outlookContact, Contact googleContact)
        {
            if (googleContact.ContactEntry.Id.Uri == null)
                throw new NullReferenceException("GoogleContact must have a valid Id");

            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;

            try
            {
                userProperties = outlookContact.UserProperties;
                prop = userProperties[sync.OutlookPropertyNameId];
                //check if outlook contact aready has google id property.
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, false);

                prop.Value = googleContact.ContactEntry.Id.Uri.Content;
            }
            catch (Exception ex)
            {
                Logger.Log(ex, EventType.Debug);
                Logger.Log("Name: " + sync.OutlookPropertyNameId, EventType.Debug);
                Logger.Log("Value: " + googleContact.ContactEntry.Id.Uri.Content, EventType.Debug);
                throw;
            }
            finally
            {
                if (prop != null)
                    Marshal.ReleaseComObject(prop);
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
            SetOutlookLastSync(sync, outlookContact);
        }

        public static void SetOutlookLastSync(ContactsSynchronizer sync, Outlook.ContactItem outlookContact)
        {
            //save sync datetime
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;
            try
            {
                userProperties = outlookContact.UserProperties;
                prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, false);
                prop.Value = DateTime.Now;
            }
            finally
            {
                if (prop != null)
                    Marshal.ReleaseComObject(prop);
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
        }

        public static DateTime? GetOutlookLastSync(ContactsSynchronizer sync, Outlook.ContactItem outlookContact)
        {
            DateTime? result = null;

            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;

            try
            {
                userProperties = outlookContact.UserProperties;
                prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop != null)
                    result = (DateTime)prop.Value;
            }
            finally
            {
                if (prop != null)
                    Marshal.ReleaseComObject(prop);
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
            return result;
        }

        public static string GetOutlookGoogleContactId(ContactsSynchronizer sync, Outlook.ContactItem outlookContact)
        {
            string id = null;
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty idProp = null;
            try
            {
                userProperties = outlookContact.UserProperties;
                idProp = userProperties[sync.OutlookPropertyNameId];
                if (idProp != null)
                {
                    id = (string)idProp.Value;
                    if (id == null)
                        throw new Exception();
                }
            }
            finally
            {
                if (idProp != null)
                    Marshal.ReleaseComObject(idProp);
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
            return id;
        }

        public static void ResetOutlookGoogleContactId(ContactsSynchronizer sync, Outlook.ContactItem outlookContact)
        {
            Outlook.UserProperties userProperties = null;

            try
            {
                userProperties = outlookContact.UserProperties;

                for (var i = userProperties.Count; i > 0; i--)
                {
                    Outlook.UserProperty p = null;
                    try
                    {
                        p = userProperties[i];
                        if (p.Name == sync.OutlookPropertyNameId || p.Name == sync.OutlookPropertyNameSynced)
                        {
                            userProperties.Remove(i);
                        }
                    }
                    finally
                    {
                        if (p != null)
                            Marshal.ReleaseComObject(p);
                    }
                }
            }
            finally
            {
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
        }

        public static string GetOutlookEmailAddress1(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email1AddressType, outlookContactItem.Email1EntryID, outlookContactItem.Email1Address);
        }

        public static string GetOutlookEmailAddress2(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email2AddressType, outlookContactItem.Email2EntryID, outlookContactItem.Email2Address);
        }

        public static string GetOutlookEmailAddress3(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email3AddressType, outlookContactItem.Email3EntryID, outlookContactItem.Email3Address);
        }

        private static string GetOutlookEmailAddress(Outlook.ContactItem outlookContactItem, string emailAddressType, string emailEntryID, string emailAddress)
        {
            switch (emailAddressType)
            {
                case "EX":  // Microsoft Exchange address: "/o=xxxx/ou=xxxx/cn=Recipients/cn=xxxx"

                    Outlook.NameSpace outlookNameSpace = null;
                    Outlook.Recipient recipient = null;
                    Outlook.AddressEntry addressEntry = null;
                    Outlook.ExchangeUser exchangeUser = null;

                    try
                    {
                        outlookNameSpace = outlookContactItem.Application.GetNamespace("mapi");
                        // The emailEntryID is garbage (bug in Outlook 2007 and before?) - so we cannot do GetAddressEntryFromID().
                        // Instead we create a temporary recipient and ask Exchange to resolve it, then get the SMTP address from it.
                        //Outlook.AddressEntry addressEntry = outlookNameSpace.GetAddressEntryFromID(emailEntryID);
                        recipient = outlookNameSpace.CreateRecipient(emailAddress);

                        recipient.Resolve();
                        if (recipient.Resolved)
                        {
                            addressEntry = recipient.AddressEntry;
                            if (addressEntry != null)
                            {
                                if (addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                                {
                                    exchangeUser = addressEntry.GetExchangeUser();
                                    if (exchangeUser != null)
                                    {
                                        return exchangeUser.PrimarySmtpAddress;
                                    }
                                }
                                else
                                {
                                    Logger.Log(string.Format("Unsupported AddressEntryUserType {0} for contact '{1}'.", addressEntry.AddressEntryUserType, outlookContactItem.FileAs), EventType.Debug);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                        // TODO: Can we do better?
                        Logger.Log(string.Format("Error getting the email address of outlook contact '{0}' from Exchange format '{1}': {2}", outlookContactItem.FileAs, emailAddress, ex.Message), EventType.Warning);
                        return emailAddress;
                    }
                    finally
                    {
                        if (exchangeUser != null)
                            Marshal.ReleaseComObject(exchangeUser);
                        if (addressEntry != null)
                            Marshal.ReleaseComObject(addressEntry);
                        if (recipient != null)
                            Marshal.ReleaseComObject(recipient);
                        if (outlookNameSpace != null)
                            Marshal.ReleaseComObject(outlookNameSpace);
                    }

                    // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                    // TODO: Can we do better?                   
                    return emailAddress;

                case "SMTP":
                default:
                    return emailAddress;
            }
        }
    }
}
