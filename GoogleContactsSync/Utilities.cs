using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Net;
using System.IO;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Collections.ObjectModel;
using Google.Contacts;

namespace GoContactSyncMod
{
    internal static class Utilities
    {
        //private static string tempPhotoPath = AppDomain.CurrentDomain.BaseDirectory + "\\TempOutlookContactPhoto.jpg";
        //private static string tempPhotoPath = Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\" + System.Windows.Forms.Application.ProductName + @"\TempOutlookContactPhoto.jpg";
        private static string tempPhotoPath = Path.GetTempPath() + @"\TempOutlookContactPhoto.jpg";


        public static byte[] BitmapToBytes(Bitmap bitmap)
        {
            //bitmap
            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
                return stream.ToArray();
            }
        }

        public static bool HasPhoto(Contact googleContact)
        {
            if (googleContact.PhotoEtag == null)
                return false;
            return true;
        }
        public static bool HasPhoto(Outlook.ContactItem outlookContact)
        {
            return outlookContact.HasPicture;
        }

        public static bool SaveGooglePhoto(Synchronizer sync, Contact googleContact, Image image)
        {
            if (googleContact.ContactEntry.PhotoUri == null)
                throw new Exception("Must reload contact from google.");

            try
            {
                using (var client = new WebClient())
                {
                    //client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.ContactsRequest.Service.QueryClientLoginToken());
                    client.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + sync.ContactsRequest.Settings.OAuth2Parameters.AccessToken);
                    client.Headers.Add(HttpRequestHeader.ContentType, "image/*");
                    using (var pic = new Bitmap(image))
                    {
                        using (var s = client.OpenWrite(googleContact.ContactEntry.PhotoUri.AbsoluteUri, "PUT"))
                        {
                            byte[] bytes = BitmapToBytes(pic);
                            s.Write(bytes, 0, bytes.Length);
                            s.Flush();
                        }
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
        public static Image GetGooglePhoto(Synchronizer sync, Contact googleContact)
        {
            if (!HasPhoto(googleContact))
                return null;

            try
            {
                using (WebClient client = new WebClient())
                {
                    //client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.ContactsRequest.Service.QueryClientLoginToken());
                    client.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + sync.ContactsRequest.Settings.OAuth2Parameters.AccessToken);
                    Stream stream = client.OpenRead(googleContact.PhotoUri.AbsoluteUri);
                    BinaryReader reader = new BinaryReader(stream);
                    Image image = Image.FromStream(stream);
                    reader.Close();

                    return image;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, string fullImagePath)
        {
            try
            {
                outlookContact.AddPicture(fullImagePath);
                //outlookContact.Save();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, Image image)
        {
            try
            {
                image.Save(tempPhotoPath);
                return SetOutlookPhoto(outlookContact, tempPhotoPath);
            }
            catch
            {
                return false;
            }
        }
        public static Image GetOutlookPhoto(Outlook.ContactItem outlookContact)
        {
            if (!HasPhoto(outlookContact))
                return null;

            try
            {
                foreach (Outlook.Attachment a in outlookContact.Attachments)
                {
                    // CH Fixed this to Contains, due to outlook picture that looks like "ContactPicture_138382.jpg"
                    if (a.DisplayName.ToUpper().Contains("CONTACTPICTURE") || a.DisplayName.ToUpper().Contains("CONTACTPHOTO"))
                    {

                        //TODO: Check why always the first added picture is returned
                        //If you add another picture, still the old picture is saved to tempPhotoPath
                        a.SaveAsFile(tempPhotoPath); 

                        using (var img = Image.FromFile(tempPhotoPath))
                        {
                            using (var bmp = new Bitmap(img))
                            {
                                return bmp;
                            }
                        }
                    }
                }
                return null;
            }
            catch
            {
                // There's an error here... If Outlook says it has a contact photo, and we can't get it, Something's broken.

                return null;
            }
        }

        public static Image CropImageGoogleFormat(Image original)
        {
            // crop image to a square in the center

            int width, height, diff;
            Point p;
            Rectangle r;

            if (original.Height == original.Width)
                return original;
            if (original.Height > original.Width)
            {
                // tall image
                width = original.Width;
                height = width;

                diff = original.Height - height;
                p = new Point(0, diff / 2);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
            else
            {
                // flat image
                height = original.Height;
                width = height;

                diff = original.Width - width;
                p = new Point(diff / 2, 0);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
        }
        public static Image CropImage(Image original, Rectangle cropArea)
        {
            using (Bitmap bmpImage = new Bitmap(original))
            {
                Bitmap bmpCrop = bmpImage.Clone(cropArea, bmpImage.PixelFormat);
                return (Image)(bmpCrop);
            }
        }

        public static void DeleteTempPhoto()
        {
            try
            {
                if (File.Exists(tempPhotoPath))
                    File.Delete(tempPhotoPath);
            }
            catch { }
        }

        public static bool ContainsGroup(Synchronizer sync, Contact googleContact, string groupName)
        {
            Group group = sync.GetGoogleGroupByName(groupName);
            if (group == null)
                return false;
            return ContainsGroup(googleContact, group);
        }
        public static bool ContainsGroup(Contact googleContact, Group group)
        {
            foreach (GroupMembership m in googleContact.GroupMembership)
            {
                if (m.HRef == group.GroupEntry.Id.AbsoluteUri)
                    return true;
            }
            return false;
        }
        public static bool ContainsGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (outlookContact.Categories == null)
                return false;

            return outlookContact.Categories.Contains(group);
        }

        public static Collection<Group> GetGoogleGroups(Synchronizer sync, Contact googleContact)
        {
            int c = googleContact.GroupMembership.Count;
            Collection<Group> groups = new Collection<Group>();
            string id;
            Group group;
            for (int i = 0; i < c; i++)
            {
                id = googleContact.GroupMembership[i].HRef;
                group = sync.GetGoogleGroupById(id);

                groups.Add(group);
            }
            return groups;
        }
        public static void AddGoogleGroup(Contact googleContact, Group group)
        {
            if (ContainsGroup(googleContact, group))
                return;

            GroupMembership m = new GroupMembership();
            m.HRef = group.GroupEntry.Id.AbsoluteUri;
            googleContact.GroupMembership.Add(m);
        }
        public static void RemoveGoogleGroup(Contact googleContact, Group group)
        {
            if (!ContainsGroup(googleContact, group))
                return;

            // TODO: broken. removes group membership but does not remove contact
            // from group in the end.

            // look for id
            GroupMembership mem;
            for (int i = 0; i < googleContact.GroupMembership.Count; i++)
            {
                mem = googleContact.GroupMembership[i];
                if (mem.HRef == group.GroupEntry.Id.AbsoluteUri)
                {
                    googleContact.GroupMembership.Remove(mem);
                    return;
                }
            }
            throw new Exception("Did not find group");
        }

        public static string[] GetOutlookGroups(string outlookContactCategories)
        {
            if (outlookContactCategories == null)
                return new string[] { };

            char[] listseparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator.ToCharArray(); 
            if (!outlookContactCategories.Contains(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator))
            {// ListSeparator doesn't work always, because ListSeparator returns "," instead of ";"
                listseparator = ",".ToCharArray();
                if (!outlookContactCategories.Contains(","))
                    listseparator = ";".ToCharArray();
            }
            string[] categories = outlookContactCategories.Split(listseparator);

            for (int i = 0; i < categories.Length; i++)
            {
                categories[i] = categories[i].Trim();
            }
            return categories;
        }
        public static void AddOutlookGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (ContainsGroup(outlookContact, group))
                return;

            // append
            if (outlookContact.Categories == null)
                outlookContact.Categories = "";
            if (!string.IsNullOrEmpty(outlookContact.Categories))
                outlookContact.Categories += ", " + group;
            else
                outlookContact.Categories += group;
        }
        public static void RemoveOutlookGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (!ContainsGroup(outlookContact, group))
                return;

            outlookContact.Categories = outlookContact.Categories.Replace(", " + group, "");
            outlookContact.Categories = outlookContact.Categories.Replace(group, "");
        }

        //ToDo: Workaround to save google Content is also not working, beause of error when closing the StreamWriter
        //public static bool SaveGoogleNoteContent(Syncronizer sync, Google.Documents.Document updated, Google.Documents.Document googleNote)
        //{

        //    if (updated.DocumentEntry.EditUri == null || googleNote.MediaSource == null)
        //        throw new Exception("Must reload note from google.");

        //    StreamWriter writer = null;
        //    StreamReader reader = null;
        //    WebClient client = null;
        //    try
        //    {
        //        client = new WebClient();
        //        client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.DocumentsRequest.Service.QueryClientLoginToken());
        //        client.Headers.Add(HttpRequestHeader.ContentType, googleNote.MediaSource.ContentType);
        //        Stream s = client.OpenWrite(updated.DocumentEntry.EditUri.ToString(), "PUT");
        //        writer = new StreamWriter(s);
        //        reader = new StreamReader(googleNote.MediaSource.GetDataStream());
        //        string body = reader.ReadToEnd();
        //        writer.Write(body);
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //    finally
        //    {
        //        if (client != null)
        //            client.Dispose();
        //        if (writer != null)
        //            writer.Close(); //This throws an exception 400 (Ungültige Anforderung)
        //        if (reader != null)
        //            reader.Close();

        //    }

        //    return true;
        //}

        static public string ConvertToText(string rtf)
        {
            using (var rtb = new System.Windows.Forms.RichTextBox())
            {
                rtb.Rtf = rtf;
                return rtb.Text;
            }
        }

        static public string ConvertToText(byte[] rtf)
        {
            string ret = null;
            if (rtf != null)
            {
                System.Text.Encoding encoding = new System.Text.ASCIIEncoding();
                ret = encoding.GetString(rtf);
                ret = ConvertToText(ret);
            }

            return ret;
        }
    }

    public class OutlookFolder : IComparable 
    {
        private string _folderName;
        private string _folderID;
        private bool _isDefaultFolder;


        public OutlookFolder(string folderName, string folderID, bool isDefaultFolder)
        {

            this._folderName = folderName;
            this._folderID   = folderID;
            this._isDefaultFolder = isDefaultFolder;
        }

        public string FolderName
        {
            get
            {
                return _folderName;
            }
        }

        public string FolderID
        {

            get
            {
                return _folderID;
            }
        }

        public bool IsDefaultFolder
        {

            get
            {
                return _isDefaultFolder;
            }
        }

        public string DisplayName
        {
            get
            {
                return _folderName + (_isDefaultFolder?" (Default)":String.Empty);
            }
        }

        public int CompareTo(object obj)
        {
            if (obj is OutlookFolder)
            {
                return this._folderName.CompareTo((obj as OutlookFolder)._folderName);
            }
            else
                throw new ArgumentException("Object is not a OutlookFolder");
        }

        public override string ToString()
        {
            if (this is OutlookFolder)
                return FolderID;
            else
                return base.ToString();
            
        }
    }

    public class GoogleCalendar : IComparable
    {
        private string _folderName;
        private string _folderID;
        private bool _isDefaultFolder;


        public GoogleCalendar(string folderName, string folderID, bool isDefaultFolder)
        {

            this._folderName = folderName;
            this._folderID = folderID;
            this._isDefaultFolder = isDefaultFolder;
        }

        public string FolderName
        {
            get
            {
                return _folderName;
            }
        }

        public string FolderID
        {

            get
            {
                return _folderID;
            }
        }

        public bool IsDefaultFolder
        {

            get
            {
                return _isDefaultFolder;
            }
        }

        public string DisplayName
        {
            get
            {
                return _folderName + (_isDefaultFolder ? " (Default)" : String.Empty);
            }
        }

        public int CompareTo(object obj)
        {
            if (obj is GoogleCalendar)
            {
                return this._folderName.CompareTo((obj as GoogleCalendar)._folderName);
            }
            else
                throw new ArgumentException("Object is not a GoogleCalendar");
        }

        public override string ToString()
        {
            if (this is GoogleCalendar)
                return FolderID;
            else
                return base.ToString();

        }
    }
}
