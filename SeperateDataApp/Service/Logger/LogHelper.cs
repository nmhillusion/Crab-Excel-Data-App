using System;
using System.Windows.Controls;

namespace SeperateDataApp.Service.Logger
{
    class LogHelper
    {
        private readonly object owner;
        private readonly LogWriteToFileHelper logWriteToFileHelper = LogWriteToFileHelper.GetInstance();
        private readonly ListViewLogHelper listViewLogHelper = ListViewLogHelper.GetInstance();

        public LogHelper(object owner)
        {
            this.owner = owner;
        }

        public void SetLogListView(ListView logListView)
        {
            this.listViewLogHelper.SetLogListView(logListView);
        }

        private void WriteLine(LogLevel logLevel, object data)
        {
            string messageLog = $" {DateTime.Now:yyyy-MM-dd HH:mm:ss.FFF} : {owner.GetType()} : [{logLevel.GetLogLevelValue()}] : {data}";
            System.Diagnostics.Debug.WriteLine(messageLog);
            logWriteToFileHelper.AppendNewLineLog(messageLog);
            listViewLogHelper.AddLogToListView(messageLog);
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
