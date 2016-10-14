using System.Collections.Generic;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod
{
    class Factory
    {

        internal static Event NewEvent()
        {
            Event ev = new Event();
            ev.Reminders = new Event.RemindersData { Overrides = new List<EventReminder>(), UseDefault = false };
            ev.Recurrence = new List<string>();
            ev.ExtendedProperties = new Event.ExtendedPropertiesData { Shared = new Dictionary<string, string>() };
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();
            ev.Locked = false;

            return ev;
        }
    }
}
