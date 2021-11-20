namespace SeperateDataApp.Service.Log
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
}