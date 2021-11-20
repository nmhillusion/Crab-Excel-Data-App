using SeperateDataApp.Service.Log;
using System;

namespace SeperateDataApp.Service
{
    class LogHelper
    {
        private readonly object owner;
        private readonly LogWriteToFileHelper logWriteToFileHelper = LogWriteToFileHelper.GetInstance();

        public LogHelper(object owner)
        {
            this.owner = owner;
        }

        private void WriteLine(LogLevel logLevel, object data)
        {
            string messageLog = $" {DateTime.Now:yyyy-MM-dd HH:mm:ss.FFF} : {owner.GetType()} : [{logLevel.GetLogLevelValue()}] : {data}";
            System.Diagnostics.Debug.WriteLine(messageLog);
            logWriteToFileHelper.AppendNewLineLog(messageLog);
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
