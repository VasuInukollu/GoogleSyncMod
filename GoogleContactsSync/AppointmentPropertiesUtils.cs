using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
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
        public static void SetOutlookGoogleAppointmentId(AppointmentsSynchronizer sync, Outlook.AppointmentItem outlookAppointment, Event googleAppointment)
        {
            if (googleAppointment.Id == null)
                throw new NullReferenceException("GoogleAppointment must have a valid Id");

            //check if outlook Appointment aready has google id property.
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;
            try
            {
                userProperties = outlookAppointment.UserProperties;
                prop = userProperties[sync.OutlookPropertyNameId];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, false);

                prop.Value = googleAppointment.Id;
            }
            catch (Exception ex)
            {
                Logger.Log(ex, EventType.Debug);
                Logger.Log("Name: " + sync.OutlookPropertyNameId, EventType.Debug);
                Logger.Log("Value: " + googleAppointment.Id, EventType.Debug);
                throw;
            }
            finally
            {
                if (prop != null)
                    Marshal.ReleaseComObject(prop);
                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
            SetOutlookLastSync(sync, outlookAppointment);
        }

        public static void SetOutlookLastSync(AppointmentsSynchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            //save sync datetime
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;
            try
            {
                userProperties = outlookAppointment.UserProperties;
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

        public static DateTime? GetOutlookLastSync(AppointmentsSynchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            DateTime? result = null;
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty prop = null;

            try
            {
                userProperties = outlookAppointment.UserProperties;
                prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop != null)
                {
                    result = (DateTime)prop.Value;
                }
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

        public static string GetOutlookGoogleAppointmentId(AppointmentsSynchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            string id = null;

            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty idProp = null;

            try
            {
                userProperties = outlookAppointment.UserProperties;
                idProp = userProperties[sync.OutlookPropertyNameId];
                if (idProp != null)
                {
                    id = (string)idProp.Value;
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

        public static void ResetOutlookGoogleAppointmentId(AppointmentsSynchronizer sync, Outlook.AppointmentItem outlookAppointment)
        {
            Outlook.UserProperties userProperties = null;

            try
            {
                userProperties = outlookAppointment.UserProperties;

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
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static string GetOutlookEmailAddress(string subject, Outlook.Recipient recipient)
        {
            string emailAddress = recipient.Address != null ? recipient.Address : recipient.Name;

            switch (recipient.AddressEntry.AddressEntryUserType)
            {
                case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:  // Microsoft Exchange address: "/o=xxxx/ou=xxxx/cn=Recipients/cn=xxxx"
                    Outlook.AddressEntry addressEntry = null;
                    Outlook.ExchangeUser exchangeUser = null;
                    try
                    {
                        // The emailEntryID is garbage (bug in Outlook 2007 and before?) - so we cannot do GetAddressEntryFromID().
                        // Instead we create a temporary recipient and ask Exchange to resolve it, then get the SMTP address from it.
                        //Outlook.AddressEntry addressEntry = outlookNameSpace.GetAddressEntryFromID(emailEntryID);

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
                                    Logger.Log(string.Format("Unsupported AddressEntryUserType {0} for email '{1}' in appointment '{2}'.", addressEntry.AddressEntryUserType, addressEntry.Address, subject), EventType.Debug);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                        // TODO: Can we do better?
                        Logger.Log(string.Format("Error getting the email address of outlook appointment '{0}' from Exchange format '{1}': {2}", subject, emailAddress, ex.Message), EventType.Warning);
                        return emailAddress;
                    }
                    finally
                    {
                        if (exchangeUser != null)
                            Marshal.ReleaseComObject(exchangeUser);
                        if (addressEntry != null)
                            Marshal.ReleaseComObject(addressEntry);
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
