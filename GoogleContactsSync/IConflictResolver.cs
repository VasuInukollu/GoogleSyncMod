using System.Collections.Generic;
using Google.Contacts;

namespace GoContactSyncMod
{
    internal interface IConflictResolver
    {
        /// <summary>
        /// Resolves contact sync conflics.
        /// </summary>
        /// <param name="outlookContact"></param>
        /// <param name="googleContact"></param>
        /// <returns>Returns ConflictResolution (enum)</returns>
        ConflictResolution Resolve(ContactMatch match, bool isNewMatch);

        ConflictResolution ResolveDuplicate(OutlookContactInfo outlookContact, List<Contact> googleContacts, out Contact googleContact);

        DeleteResolution ResolveDelete(OutlookContactInfo outlookContact);

        DeleteResolution ResolveDelete(Contact googleContact);

    }

    internal enum ConflictResolution
    {
        Skip,
        Cancel,
        OutlookWins,
        GoogleWins,
        OutlookWinsAlways,
        GoogleWinsAlways,
        SkipAlways
    }

    internal enum DeleteResolution
    {
        Cancel,
        DeleteOutlook,
        DeleteGoogle,
        KeepOutlook,
        KeepGoogle,
        DeleteOutlookAlways,
        DeleteGoogleAlways,
        KeepOutlookAlways,
        KeepGoogleAlways
    }
}
