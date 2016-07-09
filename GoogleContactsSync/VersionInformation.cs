using System;
using System.Runtime.InteropServices;
using System.Management;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.Net.Http;
//using HtmlAgilityPack;
using System.Xml.Linq;


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
     /*   public static async Task<bool> isNewVersionAvailable2()
        {

            Logger.Log("Reading version number from sf.net...", EventType.Information);
            try
            {
                //parse download site for html redirect tag
                var cookies = new CookieContainer();
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(DOWNLOADURL);
                request.UserAgent = USERAGENT;
                request.CookieContainer = cookies;
                request.AllowAutoRedirect = false;
                var response = await request.GetResponseAsync();
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream, ASCIIEncoding.ASCII);
                string strResponse = await reader.ReadToEndAsync();
                response.Close();

                //parse first html document for mirror download code
                HtmlDocument htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(strResponse);

                //var httpMetaRefresh = htmlDoc.DocumentNode.SelectNodes("//meta[@http-equiv='refresh']");
                //search for meta tag
                var xpath = "//meta[@http-equiv='refresh' and contains(@content, 'url')]";
                var refresh = htmlDoc.DocumentNode.SelectSingleNode(xpath);
                //get download url from contant
                var content = refresh.Attributes["content"].Value;
                //extract url
                var secondDownloadUrl = Regex.Match(content, @"\s*url\s*=\s*([^ ]+)").Groups[1].Value.Trim();
              
                //check sf.net site for next redirect 
                request = (HttpWebRequest)WebRequest.Create(secondDownloadUrl);
                request.UserAgent = USERAGENT;
                request.CookieContainer = cookies;
                request.AllowAutoRedirect = true;
                response = (HttpWebResponse)request.GetResponse();
                
                //extracting version number from url
                const string firstPattern = "Releases/";
                // ex. /project/googlesyncmod/Releases/3.9.5/SetupGCSM-3.9.5.msi
                string webVersion = response.ResponseUri.AbsolutePath;
                response.Close();

                //get version number string
                int first = webVersion.IndexOf(firstPattern) + firstPattern.Length;
                int second = webVersion.IndexOf("/", first);
                Version webVersionNumber = new Version(webVersion.Substring(first, second - first));

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
            catch (Exception ex)
            {
                Logger.Log("Could not read version number from sf.net...", EventType.Information);
                Logger.Log(ex, EventType.Debug);
                return false;
            }
        }
   */ 
        public static async Task<bool> isNewVersionAvailable()
        {
            Logger.Log("Reading version number from sf.net...", EventType.Information);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var response = await client.GetAsync("https://sourceforge.net/projects/googlesyncmod/files/updates_v1.xml", HttpCompletionOption.ResponseHeadersRead);
                    response.EnsureSuccessStatusCode();
                    var stream = await response.Content.ReadAsStreamAsync();
                    var doc = XDocument.Load(stream);
                    
                    var webVersionNumber = new Version(doc.Element("Version").Value);
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
