using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace SeperateDataApp.Service.Logger
{
    class ListViewLogHelper
    {
        private static readonly ListViewLogHelper instance = new();
        private readonly int MAX_LOG_LINES = 50;
        private ListView logListView;

        private ListViewLogHelper()
        {
        }

        public static ListViewLogHelper GetInstance()
        {
            return instance;
        }

        public void SetLogListView(ListView _logListView)
        {
            logListView = _logListView;
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public void AddLogToListView(string messageLog)
        {
            if (null != logListView)
            {
                logListView.Dispatcher.InvokeAsync(() =>
                {
                    if (MAX_LOG_LINES < logListView.Items.Count)
                    {
                        logListView.Items.RemoveAt(0);
                    }

                    logListView.Items.Add(
                        messageLog
                    );

                    logListView.ScrollIntoView(messageLog);
                });
            }
        }
    }
}
