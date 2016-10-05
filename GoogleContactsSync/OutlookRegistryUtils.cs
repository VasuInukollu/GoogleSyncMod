using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;

namespace GoContactSyncMod
{
    internal static class OutlookRegistryUtils
    {
        public static string GetOutlookVersion()
        {
            string ret = string.Empty;
            RegistryKey registryOutlookKey = null;

            try
            {
                int outlookVersion = GetMajorVersion(GetOutlookPath());
                ret = GetMajorVersionToString(outlookVersion);

                string outlookKey = @"Software\Wow6432Node\Microsoft\Office\" + outlookVersion + @".0\Outlook";
                registryOutlookKey = Registry.LocalMachine.OpenSubKey(outlookKey, false);
                if (registryOutlookKey == null)
                {
                    outlookKey = @"Software\Microsoft\Office\" + outlookVersion + @".0\Outlook";
                    registryOutlookKey = Registry.LocalMachine.OpenSubKey(outlookKey, false);
                }
                if (registryOutlookKey != null)
                {
                    string bitness = registryOutlookKey.GetValue(@"Bitness", @" (unknown bitness)").ToString();
                    if (string.IsNullOrEmpty(bitness))
                    {
                        return ret + @" (unknown bitness)";
                    }
                    else
                    {
                        if (bitness == @"x86")
                            return ret + @" (32-bit)";
                        else if (bitness == @"x64")
                            return ret + @" (64-bit)";
                        else
                            return ret + @" (unknown)";
                    }
                }
            }
            catch
            {
            }
            finally
            {
                if (registryOutlookKey != null)
                    registryOutlookKey.Close();
            }

            return ret;
        }

        public static string GetPossibleErrorDiagnosis()
        {
            int outlookVersion = GetMajorVersion(GetOutlookPath());
            string diagnosis = CheckOfficeRegistry(outlookVersion);
            return "Could not connect to 'Microsoft Outlook'.\r\nYou have " + GetMajorVersionToString(outlookVersion) + " installed.\r\n" + diagnosis;
        }

        private static string CheckOfficeRegistry(int outlookVersion)
        {
            string toReturn = string.Empty;
            string registryVersion = ConvertMajorVersionToRegistryVersion(outlookVersion);

            const string interfaceVersion = @"Interface\{00063001-0000-0000-C000-000000000046}\TypeLib";
            RegistryKey interfaceKey = Registry.ClassesRoot.OpenSubKey(interfaceVersion, false);
            if (interfaceKey != null)
            {
                String typeLib = interfaceKey.GetValue(string.Empty).ToString();
                if (typeLib != "{00062FFF-0000-0000-C000-000000000046}")
                    return "Your registry " + interfaceKey.ToString() + " points to TypeLib " + typeLib + " and should to {00062FFF-0000-0000-C000-000000000046}" + registryVersion + ".\r\nPlease read FAQ and fix your Office installation";
                String version = interfaceKey.GetValue("Version").ToString();
                if (version != registryVersion)
                    return "Your registry " + interfaceKey.ToString() + " points to version " + version + " and your Outlook is installed with version " + registryVersion + ".\r\nPlease read FAQ and fix your Office installation";
            }
            else
            {
                return "Cannot open registry " + interfaceKey.ToString() + ".\r\nPlease read FAQ and fix your Office installation";
            }

            if (!string.IsNullOrEmpty(registryVersion))
            {
                string RegKey = @"TypeLib\{00062FFF-0000-0000-C000-000000000046}\" + registryVersion + @"\0\";

                RegistryKey mainKey = Registry.ClassesRoot.OpenSubKey(RegKey + "win32", false);
                if (mainKey != null)
                {
                    String path = mainKey.GetValue(string.Empty).ToString();
                    if (!File.Exists(path))
                        return "Your registry " + mainKey.ToString() + " points to file " + path + " and this file does not exist.\r\nPlease read FAQ and fix your Office installation";
                }
                mainKey = Registry.ClassesRoot.OpenSubKey(RegKey + "win64", false);
                if (mainKey != null)
                {
                    String path = mainKey.GetValue(string.Empty).ToString();
                    if (!File.Exists(path))
                        return "Your registry " + mainKey.ToString() + " points to file " + path + " and this file does not exist.\r\nPlease read FAQ and fix your Office installation";
                }

                mainKey = Registry.ClassesRoot.OpenSubKey(@"TypeLib\{00062FFF-0000-0000-C000-000000000046}\", false);
                String[] keys = mainKey.GetSubKeyNames();
                if (keys.Length > 1)
                {
                    string allKeys = "";
                    for (int i = 0; i < keys.Length; i++)
                    {
                        string element = keys[i];
                        if (element != registryVersion)
                        {
                            allKeys = allKeys + element + ",";
                        }
                    }
                    allKeys = allKeys.Substring(0, allKeys.Length - 1);
                    return "Your registry " + mainKey.ToString() + " points to Office versions: " + allKeys + " other than you have installed: " + registryVersion + ".\r\nPlease read FAQ and fix your Office installation";
                }
            }
            return toReturn;
        }

        private static string GetOutlookPath()
        {
            const string regKey = @"Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe";
            string toReturn = string.Empty;

            try
            {
                RegistryKey mainKey = Registry.CurrentUser.OpenSubKey(regKey, false);
                if (mainKey != null)
                {
                    toReturn = mainKey.GetValue(string.Empty).ToString();
                }
            }
            catch
            { }

            if (string.IsNullOrEmpty(toReturn))
            {
                try
                {
                    RegistryKey mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                    if (mainKey != null)
                    {
                        toReturn = mainKey.GetValue(string.Empty).ToString();
                    }
                }
                catch
                { }
            }

            return toReturn;
        }

        private static int GetMajorVersion(string path)
        {
            int toReturn = 0;
            if (File.Exists(path))
            {
                try
                {
                    toReturn = FileVersionInfo.GetVersionInfo(path).FileMajorPart;
                }
                catch
                { }
            }
            return toReturn;
        }

        private static string ConvertMajorVersionToRegistryVersion(int version)
        {
            switch (version)
            {
                case 9: return "9.0";
                case 10: return "9.1";
                case 11: return "9.2";
                case 12: return "9.3";
                case 14: return "9.4";
                case 15: return "9.5";
                case 16: return "9.6";
                default: return string.Empty;
            }
        }

        private static string GetMajorVersionToString(int version)
        {
            switch (version)
            {
                case 7: return "Office 97";
                case 8: return "Office 98";
                case 9: return "Office 2000";
                case 10: return "Office XP";
                case 11: return "Office 2003";
                case 12: return "Office 2007";
                case 14: return "Office 2010";
                case 15: return "Office 2013";
                case 16: return "Office 2016";
                default: return "unknown Office version (" + version + ")";
            }
        }
    }
}
