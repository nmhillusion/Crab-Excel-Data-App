using System;

namespace SeperateDataApp.Service
{
    internal class LogLevelBase
    {
        private string logLevelValue;

        public string GetLogLevelValue()
        {
            return logLevelValue;
        }

        public LogLevelBase(string logLevelValue)
        {
            this.logLevelValue = logLevelValue;
        }
    }

    class LogLevel : LogLevelBase
    {
        public static readonly LogLevel DEBUG = new("DEBUG");
        public static readonly LogLevel INFO = new("INFO");
        public static readonly LogLevel WARN = new("WARN");
        public static readonly LogLevel ERROR = new("ERROR");

        private LogLevel(string logLevelValue) : base(logLevelValue) { }
    }

    class LogHelper
    {
        private readonly object owner;
        public LogHelper(object owner)
        {
            this.owner = owner;
        }

        private void WriteLine(LogLevel logLevel, object data)
        {
            System.Diagnostics.Debug.WriteLine($" {DateTime.Now:yyyy-MM-dd HH:mm:ss.FFF} : {owner.GetType()} : [{logLevel.GetLogLevelValue()}] : {data}");
        }

        public void Debug(object data)
        {
            WriteLine(LogLevel.DEBUG, data);
        }

        public void Info(object data)
        {
            WriteLine(LogLevel.INFO, data);
        }

        public void Warn(object data)
        {
            WriteLine(LogLevel.WARN, data);
        }

        public void Error(object data)
        {
            WriteLine(LogLevel.ERROR, data);
        }
    }
}
