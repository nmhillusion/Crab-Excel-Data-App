namespace CrabExcelDataApp.Service.Logger
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
        public static readonly LogLevel DEBUG = new LogLevel("DEBUG");
        public static readonly LogLevel INFO = new LogLevel("INFO");
        public static readonly LogLevel WARN = new LogLevel("WARN");
        public static readonly LogLevel ERROR = new LogLevel("ERROR");

        private LogLevel(string logLevelValue) : base(logLevelValue) { }
    }
}