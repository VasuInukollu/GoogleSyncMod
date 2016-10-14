using System;
using System.Runtime.InteropServices;
using System.Management;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.Net.Http;
using System.Xml.Linq;
using System.Threading;

namespace GoContactSyncMod
{
    static class VersionInformation
    {
        private const string DOWNLOADURL = "https://sourceforge.net/projects/googlesyncmod/files/latest/download";
        private const string USERAGENT = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0";

        public enum OutlookMainVersion
        {
            Outlook2002,
            Outlook2003,
            Outlook2007,
            Outlook2010,
            Outlook2013,
            Outlook2016,
            OutlookUnknownVersion,
            OutlookNoInstance
        }

        public static OutlookMainVersion GetOutlookVersion(Microsoft.Office.Interop.Outlook.Application appVersion)
        {
            if (appVersion == null)
                appVersion = new Microsoft.Office.Interop.Outlook.Application();

            switch (appVersion.Version.ToString().Substring(0, 2))
            {
                case "10":
                    return OutlookMainVersion.Outlook2002;
                case "11":
                    return OutlookMainVersion.Outlook2003;
                case "12":
                    return OutlookMainVersion.Outlook2007;
                case "14":
                    return OutlookMainVersion.Outlook2010;
                case "15":
                    return OutlookMainVersion.Outlook2013;
                case "16":
                    return OutlookMainVersion.Outlook2016;
                default:
                    {
                        if (appVersion != null)
                        {
                            Marshal.ReleaseComObject(appVersion);
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        return OutlookMainVersion.OutlookUnknownVersion;
                    }
            }

        }

        /// <summary>
        /// detect windows main version
        /// </summary>
        public static string GetWindowsVersion()
        {
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2",
                        "SELECT Caption, OSArchitecture, Version FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject managementObject in searcher.Get())
                    {
                        string versionString = managementObject["Caption"].ToString() + " (" +
                                               managementObject["OSArchitecture"].ToString() + "; " +
                                               managementObject["Version"].ToString() + ")";
                        return versionString;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, EventType.Debug);
            }

            return "Unknown Windows Version";
        }

        public static Version getGCSMVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            Version assemblyVersionNumber = new Version(fvi.FileVersion);

            return assemblyVersionNumber;
        }

        public static async Task<bool> isNewVersionAvailable(CancellationToken cancellationToken)
        {
            Logger.Log("Reading version number from sf.net...", EventType.Information);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var response = await client.GetAsync("https://sourceforge.net/projects/googlesyncmod/files/updates_v1.xml", HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                    response.EnsureSuccessStatusCode();
                    var stream = await response.Content.ReadAsStreamAsync();
                    var doc = XDocument.Load(stream);

                    var strVersion = doc.Element("Version").Value;
                    if (!string.IsNullOrEmpty(strVersion))
                    {
                        var webVersionNumber = new Version(strVersion);
                        //compare both versions
                        var result = webVersionNumber.CompareTo(getGCSMVersion());
                        if (result > 0)
                        {   //newer version found
                            Logger.Log("New version of GCSM detected on sf.net!", EventType.Information);
                            return true;
                        }
                        else
                        {   //older or same version found
                            Logger.Log("Version of GCSM is uptodate.", EventType.Information);
                            return false;
                        }
                    }
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Could not read version number from sf.net...", EventType.Information);
                Logger.Log(ex, EventType.Debug);
                return false;
            }
        }
    }
}
