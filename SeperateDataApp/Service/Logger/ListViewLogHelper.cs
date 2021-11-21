using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace SeperateDataApp.Service.Logger
{
    class ListViewLogHelper
    {
        private static readonly ListViewLogHelper instance = new();
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
                logListView.Items.Add(
                    messageLog
                );
            }
        }
    }
}
