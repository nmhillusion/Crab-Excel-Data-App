using System.IO;

namespace SeperateDataApp.Service.Log
{
    internal class LogWriteToFileHelper
    {
        private static readonly LogWriteToFileHelper instance = new();
        private readonly string fileLogPath = "./app.log";

        private LogWriteToFileHelper()
        {
        }

        public static LogWriteToFileHelper GetInstance()
        {
            return instance;
        }

        public void AppendNewLineLog(string message)
        {
            using StreamWriter streamWriter = File.AppendText(fileLogPath);
            streamWriter.WriteLine(message);
            streamWriter.Flush();
            streamWriter.Close();
        }
    }
}
