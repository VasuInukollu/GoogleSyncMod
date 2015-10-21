using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Management;
using System.Management.Instrumentation;
using System.Net;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.Threading;
using System.Globalization;

namespace GoContactSyncMod
{
    static class VersionInformation
    {
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
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", 
                    "SELECT * FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject managementObject in searcher.Get())
                {
                    string versionString = managementObject["Caption"].ToString() + " (" +
                                           managementObject["OSArchitecture"].ToString() + "; " +
                                           managementObject["Version"].ToString() + ")";
                    return versionString;
                }
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

        /// <summary>
        /// getting the newest availible version on sourceforge.net of GCSM
        /// </summary>
        public static bool isNewVersionAvailable()
        {

            Logger.Log("Reading version number from sf.net...", EventType.Information);
            try
            {
                //check sf.net site for version number
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
                request.AllowAutoRedirect = true;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                request.Abort();

                //extracting version number from url
                const string firstPattern = "Releases/";
                // ex. /project/googlesyncmod/Releases/3.9.5/SetupGCSM-3.9.5.msi
                string webVersion = response.ResponseUri.AbsolutePath;

                //get version number string
                int first = webVersion.IndexOf(firstPattern) + firstPattern.Length;
                int second = webVersion.IndexOf("/", first);
                Version webVersionNumber = new Version(webVersion.Substring(first, second - first));

                response.Close();

                //compare both versions
                var result = webVersionNumber.CompareTo(getGCSMVersion());
                if (result > 0)
                {   //newer version found
                    Logger.Log("New version of GCSM detected on sf.net!", EventType.Information);              
                    return true;
                }
                else
                {            //older or same version found
                    Logger.Log("Version of GCSM is uptodate.", EventType.Information);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Could not read version number from sf.net...", EventType.Warning);
                Logger.Log(ex.ToString(), EventType.Debug);
                return false;
            }
        }
    }
}
