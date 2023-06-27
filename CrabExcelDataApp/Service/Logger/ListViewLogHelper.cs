using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Threading;

namespace CrabExcelDataApp.Service.Logger
{
    class ListViewLogHelper
    {
        private readonly int MAX_LOG_LINES = 50;
        private ListView logListView;

        public ListViewLogHelper()
        {
        }

        public void SetLogListView(ListView _logListView)
        {
            logListView = _logListView;
        }

        //[MethodImpl(MethodImplOptions.Synchronized)]
        public void AddLogToListView(string messageLog)
        {
            Debug.WriteLine($"[pre] insert log to list view container -> {messageLog}");
            logListView?.Dispatcher.Invoke(DispatcherPriority.Normal, new ThreadStart(delegate
                {
                    Debug.WriteLine($"insert log to list view container: {logListView} -> {messageLog}");
                    if (MAX_LOG_LINES < logListView.Items.Count)
                    {
                        logListView.Items.RemoveAt(0);
                    }

                    logListView.Items.Add(
                        messageLog
                    );

                    logListView.ScrollIntoView(messageLog);
                    Debug.WriteLine("<< insert log to list view container");
                })
            );
        }
    }
}
