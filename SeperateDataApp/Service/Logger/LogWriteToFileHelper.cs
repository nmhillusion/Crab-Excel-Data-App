using System.IO;
using System.Runtime.CompilerServices;

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

        [MethodImpl(MethodImplOptions.Synchronized)]
        public void AppendNewLineLog(string message)
        {
            using StreamWriter streamWriter = File.AppendText(fileLogPath);
            streamWriter.WriteLine(message);
            streamWriter.Flush();
            streamWriter.Close();
        }
    }
}
