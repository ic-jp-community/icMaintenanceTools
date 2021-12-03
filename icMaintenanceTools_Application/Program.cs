using ICApiAddin.icMaintenanceTools;
using System;
using System.Windows.Forms;

namespace icMaintenanceTools_Application
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new icMaintenanceToolsMain());
        }
    }
}
