using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace GoContactSyncMod
{
    enum EventType
    {
        Debug,
        Information,
        Warning,
        Error
    }

    struct LogEntry
    {
        public DateTime date;
        public EventType type;
        public string msg;

        public LogEntry(DateTime _date, EventType _type, string _msg)
        {
            date = _date; type = _type;  msg = _msg;
        }
    }

    static class Logger
    {
		public static List<LogEntry> messages = new List<LogEntry>();
		public delegate void LogUpdatedHandler(string Message);
        public static event LogUpdatedHandler LogUpdated;
        private static StreamWriter logwriter = InitializeLogWriter();

        public static readonly string Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
        public static readonly string AuthFolder = Folder + "\\Auth\\";

        static StreamWriter InitializeLogWriter()
        {
            StreamWriter writer = null;

            if (!Directory.Exists(Folder))
            {
                Directory.CreateDirectory(Folder);
                Directory.CreateDirectory(AuthFolder);
            }
            try
            {
                string logFileName = Folder + "log.txt";

                //If log file is bigger than 1 MB, move it to backup file and create new file
                FileInfo logFile = new FileInfo(logFileName);
                if (logFile.Exists && logFile.Length >= 1000000)
                    File.Move(logFileName, logFileName + "_" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss"));

                writer = new StreamWriter(logFileName, true);

                writer.WriteLine("[Start Rolling]");
                writer.Flush();

                return writer;
            }
            catch (Exception ex)
            {
                if (writer != null) writer.Dispose();
                ErrorHandler.Handle(ex);
                return null;
            }
        }

        public static void Close()
        {
            try
            {
                if (logwriter != null)
                {
                    logwriter.WriteLine("[End Rolling]");
                    logwriter.Flush();
                    logwriter.Close();
                }
            }
            catch(Exception e)
            {
                ErrorHandler.Handle(e);
            }
        }

        private static string formatMessage(string message, EventType eventType)
        {
            return String.Format("{0}:{1}{2}", eventType, Environment.NewLine, message);
        }

		private static string GetLogLine(LogEntry entry)
        {
            return String.Format("[{0} | {1}]\t{2}\r\n", entry.date, entry.type, entry.msg);
        }

		public static void Log(string message, EventType eventType)
        {
            
            
            LogEntry new_logEntry = new LogEntry(DateTime.Now, eventType, message);
            messages.Add(new_logEntry);

            try
            {
                logwriter.Write(GetLogLine(new_logEntry));
                logwriter.Flush();
            }
            catch (Exception)
            {
                //ignore it, because if you handle this error, the handler will again log the message
                //ErrorHandler.Handle(ex);
            }

            //Populate LogMessage to all subscribed Logger-Outputs, but only if not Debug message, Debug messages are only logged to logfile
            if (LogUpdated != null && eventType > EventType.Debug)
                LogUpdated(GetLogLine(new_logEntry));



        }

        public static void Log(Exception ex, EventType eventType)
        {
            CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            if (ex.InnerException != null)
            {
                Log("Inner Exception Type: " + ex.InnerException.GetType().ToString(), eventType);
                Log("Inner Exception: " + ex.InnerException.Message, eventType);
                Log("Inner Source: " + ex.InnerException.Source, eventType);
                if (ex.InnerException.StackTrace != null)
                {
                    Log("Inner Stack Trace: " + ex.InnerException.StackTrace, eventType);
                }
            }
            Log("Exception Type: " + ex.GetType().ToString(), eventType);
            Log("Exception: " + ex.Message, eventType);
            Log("Source: " + ex.Source, eventType);
            if (ex.StackTrace != null)
            {
                Log("Stack Trace: " + ex.StackTrace, eventType);
            }

            Thread.CurrentThread.CurrentCulture = oldCI;
            Thread.CurrentThread.CurrentUICulture = oldCI;
        }

        /*
        public void LogUnique(string message, EventType eventType)
        {
            string logMessage = formatMessage(message, eventType);
            if (!messages.ContainsValue(logMessage))
                messages.Add(DateTime.Now, logMessage); //TODO: Outdated, no dictionary used anymore.
        }
        */

		public static void ClearLog()
        {
            messages.Clear();
        }

        /*
        public string GetText()
        {
            StringBuilder output = new StringBuilder();
            foreach (var logitem in messages)
                output.AppendLine(GetLogLine(logitem));

            return output.ToString();
        }
        */
    }
}