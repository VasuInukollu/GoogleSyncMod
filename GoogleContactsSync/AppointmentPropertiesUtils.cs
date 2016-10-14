using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections;
using System.Runtime.InteropServices;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod
{
    internal static class AppointmentPropertiesUtils
    {
        public static string GetOutlookId(Outlook.AppointmentItem outlookAppointment)
        {
            return outlookAppointment.EntryID;
        }
        public static string GetGoogleId(Event googleAppointment)
        {
            string id = googleAppointment.Id.ToString();
            if (id == null)
                throw new Exception();
            return id;
        }

        public static void SetGoogleOutlookAppointmentId(string syncProfile, Event googleAppointment, Outlook.AppointmentItem outlookAppointment)
        {
            if (outlookAppointment.EntryID == null)
                throw new Exception("Must save outlook Appointment before getting id");

            SetGoogleOutlookAppointmentId(syncProfile, googleAppointment, GetOutlookId(outlookAppointment));
        }

        public static void SetGoogleOutlookAppointmentId(string syncProfile, Event googleAppointment, string outlookAppointmentId)
        {
            // check if exists
            bool found = false;
            if (googleAppointment.ExtendedProperties == null)
            {
                googleAppointment.ExtendedProperties = new Event.ExtendedPropertiesData();
            }
            if (googleAppointment.ExtendedProperties.Shared == null)
            {
                googleAppointment.ExtendedProperties.Shared = new Dictionary<string, string>();
            }
            foreach (var p in googleAppointment.ExtendedProperties.Shared)
            {
                if (p.Key == "gos:oid:" + syncProfile + "")
                {
                    googleAppointment.ExtendedProperties.Shared[p.Key] = outlookAppointmentId;
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                var prop = new KeyValuePair<string, string>("gos:oid:" + syncProfile + "", outlookAppointmentId);
                googleAppointment.ExtendedProperties.Shared.Add(prop);
            }
        }
        public static string GetGoogleOutlookAppointmentId(string syncProfile, Event googleAppointment)
        {
            // get extended prop
            if (googleAppointment.ExtendedProperties != null && googleAppointment.ExtendedProperties.Shared != null)
            {
                foreach (var p in googleAppointment.ExtendedProperties.Shared)
                {
                    if (p.Key == "gos:oid:" + syncProfile + "")
                        return p.Value;
                }
            }
            return null;
        }
        public static void ResetGoogleOutlookAppointmentId(string syncProfile, Event googleAppointment)
        {
            if (googleAppointment.ExtendedProperties != null && googleAppointment.ExtendedProperties.Shared != null)
            {
                // get extended prop
                foreach (var p in googleAppointment.ExtendedProperties.Shared)
                {
                    if (p.Key == "gos:oid:" + syncProfile + "")
                    {
                        // remove 
                        googleAppointment.ExtendedProperties.Shared.Remove(p);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Sets the syncId of the Outlook Appointment and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="outlookAppointment"></param>
        /// <param name="googleAppointment"></param>
        public static void SetOutlookGoogleAppointmentId(Synchronizer sync, Outlook.AppointmentItem outlookAppointment, Event googleAppointment)
        {
            if (googleAppointment.Id == null)
                throw new NullReferenceException("GoogleAppointment must have a valid Id");

            //check if outlook Appointment aready has google id property.
            Outlook.UserProperties userProperties = outlookAppointment.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameId];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, true);
                try
                {
                    prop.Value = googleAppointment.Id;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }

            //save last google's updated date as property
            /*prop = outlookAppointment.UserProperties[OutlookPropertyNameUpdated];
            if (prop == null)
                prop = outlookAppointment.UserProperties.Add(OutlookPropertyNameUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = googleAppointment.Updated;*/

            //Also set the OutlookLastSync date when setting a match between Outlook and Google to assure the lastSync updated when Outlook Appointment is saved afterwards
            SetOutlookLastSync(sync, outlookAppointment);
        }

        public static void SetOutlookLastSync(Synchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            //save sync datetime
            Outlook.UserProperties userProperties = outlookAppointment.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, true);
                try
                {
                    prop.Value = DateTime.Now;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static DateTime? GetOutlookLastSync(Synchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            DateTime? result = null;
            Outlook.UserProperties userProperties = outlookAppointment.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop != null)
                {
                    try
                    {
                        result = (DateTime)prop.Value;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(prop);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return result;
        }
        public static string GetOutlookGoogleAppointmentId(Synchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            string id = null;
            Outlook.UserProperties userProperties = outlookAppointment.UserProperties;
            try
            {
                Outlook.UserProperty idProp = userProperties[sync.OutlookPropertyNameId];
                if (idProp != null)
                {
                    try
                    {
                        id = (string)idProp.Value;
                        if (id == null)
                            throw new Exception();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(idProp);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return id;
        }
        public static void ResetOutlookGoogleAppointmentId(Synchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            Outlook.UserProperties userProperties = outlookAppointment.UserProperties;
            try
            {
                Outlook.UserProperty idProp = userProperties[sync.OutlookPropertyNameId];
                try
                {
                    Outlook.UserProperty lastSyncProp = userProperties[sync.OutlookPropertyNameSynced];
                    try
                    {
                        if (idProp == null && lastSyncProp == null)
                            return;

                        List<int> indexesToBeRemoved = new List<int>();
                        IEnumerator en = userProperties.GetEnumerator();
                        en.Reset();
                        int index = 1; // 1 based collection            
                        while (en.MoveNext())
                        {
                            Outlook.UserProperty userProperty = en.Current as Outlook.UserProperty;
                            if (userProperty == idProp || userProperty == lastSyncProp)
                            {
                                indexesToBeRemoved.Add(index);
                                //outlookAppointment.UserProperties.Remove(index);
                                //Don't return to remove both properties, googleId and lastSynced
                                //return;
                            }
                            index++;
                            Marshal.ReleaseComObject(userProperty);
                        }

                        for (int i = indexesToBeRemoved.Count - 1; i >= 0; i--)
                            userProperties.Remove(indexesToBeRemoved[i]);
                        //throw new Exception("Did not find prop.");
                    }
                    finally
                    {
                        if (lastSyncProp != null)
                            Marshal.ReleaseComObject(lastSyncProp);
                    }
                }
                finally
                {
                    if (idProp != null)
                        Marshal.ReleaseComObject(idProp);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static string GetOutlookEmailAddress(string subject, Outlook.Recipient recipient)
        {
            string emailAddress = recipient.Address != null ? recipient.Address : recipient.Name;
            switch (recipient.AddressEntry.AddressEntryUserType)
            {
                case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:  // Microsoft Exchange address: "/o=xxxx/ou=xxxx/cn=Recipients/cn=xxxx"

                    try
                    {
                        // The emailEntryID is garbage (bug in Outlook 2007 and before?) - so we cannot do GetAddressEntryFromID().
                        // Instead we create a temporary recipient and ask Exchange to resolve it, then get the SMTP address from it.
                        //Outlook.AddressEntry addressEntry = outlookNameSpace.GetAddressEntryFromID(emailEntryID);

                        //try
                        //{
                        recipient.Resolve();
                        if (recipient.Resolved)
                        {
                            Outlook.AddressEntry addressEntry = recipient.AddressEntry;
                            if (addressEntry != null)
                            {
                                try
                                {
                                    if (addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                                    {
                                        Outlook.ExchangeUser exchangeUser = addressEntry.GetExchangeUser();
                                        if (exchangeUser != null)
                                        {
                                            try
                                            {
                                                return exchangeUser.PrimarySmtpAddress;
                                            }
                                            finally
                                            {
                                                Marshal.ReleaseComObject(exchangeUser);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log(string.Format("Unsupported AddressEntryUserType {0} for email '{1}' in appointment '{2}'.", addressEntry.AddressEntryUserType, addressEntry.Address, subject), EventType.Debug);
                                    }
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(addressEntry);
                                }
                            }
                        }
                        //}
                        //finally
                        //{
                        //    if (recipient != null)
                        //        Marshal.ReleaseComObject(recipient);
                        //}
                    }
                    catch (Exception ex)
                    {
                        // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                        // TODO: Can we do better?
                        Logger.Log(string.Format("Error getting the email address of outlook appointment '{0}' from Exchange format '{1}': {2}", subject, emailAddress, ex.Message), EventType.Warning);
                        return emailAddress;
                    }

                    // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                    // TODO: Can we do better?                   
                    return emailAddress;

                case Outlook.OlAddressEntryUserType.olSmtpAddressEntry:
                default:
                    return emailAddress;
            }
        }


    }
}
