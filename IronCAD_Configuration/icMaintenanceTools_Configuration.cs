using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ICApiAddin.icMaintenanceTools;
namespace IronCAD_Configuration
{
    [System.ComponentModel.RunInstaller(true)]
    public class icMaintenanceTools_Config : System.Configuration.Install.Installer
    {
        public class WindowWrapper : IWin32Window
        {
            private readonly IntPtr hwnd;
            public IntPtr Handle
            {
                get { return hwnd; }
            }
            public WindowWrapper(IntPtr handle)
            {
                hwnd = handle;
            }
        }
        public static class TopMostMessageBox
        {
            public static void Show(string message)
            {
                var proc = Process.GetProcessesByName("msiexec").FirstOrDefault(p => p.MainWindowTitle == "icMaintenanceTools");
                if (proc != null)
                {
                    MessageBox.Show(new WindowWrapper(proc.MainWindowHandle), message); // 結果、メッセージボックスも最前面に表示される
                }
                else
                {
                    MessageBox.Show(message); // 結果、メッセージボックスも最前面に表示される
                }
            }
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);
            bool ret1 = AddConfig(@"C:\Program Files\IronCAD\2020\Config\Ironcad.Addin.config",
                Addin.ADDIN_GUID,
                "icMaintenanceTools",
                "IRONCADの修復や設定変更を行うツールです。");
            bool ret2 = AddConfig(@"C:\Program Files\IronCAD\2021\Config\Ironcad.Addin.config",
                Addin.ADDIN_GUID,
                "icMaintenanceTools",
                "icMaintenanceToolsプロジェクト");
            if ((ret1 != true) && (ret2 != true))
            {
                TopMostMessageBox.Show("IRONCADが見つかりませんでした。\n(インストールは続行されます)");
            }
        }

        public override void Commit(System.Collections.IDictionary savedState)
        {
            base.Commit(savedState);
        }

        public override void Rollback(System.Collections.IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            base.Uninstall(savedState);
            for (int count = 0; count < 10; count++)
            {
                int delnum =  DeleteConfig(@"C:\Program Files\IronCAD\2020\Config\Ironcad.Addin.config",
                                        Addin.ADDIN_GUID);
                delnum +=  DeleteConfig(@"C:\Program Files\IronCAD\2021\Config\Ironcad.Addin.config",
                                        Addin.ADDIN_GUID);
                if(delnum == 0)
                {
                    break;
                }
            }
        }

        public static bool AddConfig(string xmlpath, string guid, string name, string description)
        {
            /* ファイルの有無チェック */
            if (File.Exists(xmlpath) != true)
            {
                return false;
            }

            //C:\test.txt の属性を取得する
            System.IO.FileAttributes attr = System.IO.File.GetAttributes(xmlpath);

            //読み取り専用属性があるか調べる
            if ((attr & System.IO.FileAttributes.ReadOnly) ==
                System.IO.FileAttributes.ReadOnly)
            {
                //読み取り専用属性を削除する
                System.IO.File.SetAttributes(xmlpath, attr & (~System.IO.FileAttributes.ReadOnly));
            }

            StreamReader sr = new StreamReader(xmlpath);
            List<string> allLine = new List<string>();
            while (sr.EndOfStream != true)
            {
                allLine.Add(sr.ReadLine());
            }
            sr.Close();

            int endIndex = -1;
            for (int i = allLine.Count - 1; i >= 0; i--)
            {
                string str = allLine[i].Trim(' ', '\t');
                if (str.Contains("</AddIns>") == true)
                {
                    endIndex = i;
                    break;
                }
            }
            int addSpaceNum = allLine[endIndex - 1].IndexOf('<');
            string space = string.Empty;
            for (int i = 0; i < addSpaceNum; i++)
            {
                space += " ";
            }
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(space + "<AddIn>");
            sb.AppendLine(space + " <identify></identify>");
            sb.AppendLine(space + string.Format(" <inprocserver>{0}</inprocserver>", name));
            sb.AppendLine(space + string.Format(" <siteclsid>{0}</siteclsid>", "{" + guid + "}"));
            sb.AppendLine(space + " <modulever></modulever>");
            sb.AppendLine(space + " <platformver></platformver>");
            sb.AppendLine(space + " <autoload></autoload>");
            sb.AppendLine(space + string.Format(" <system>{0}</system>", true));
            sb.AppendLine(space + " <displayname>");
            sb.AppendLine(space + string.Format("   <default>{0}</default>", name));
            sb.AppendLine(space + "   <zh-cn></zh-cn>");
            sb.AppendLine(space + " </displayname>");
            sb.AppendLine(space + " <description>");
            sb.AppendLine(space + string.Format("   <default>{0}</default>", description));
            sb.AppendLine(space + "   <zh-cn></zh-cn>");
            sb.AppendLine(space + " </description>");
            sb.AppendLine(space + "</AddIn>");

            StreamWriter sw = new StreamWriter(xmlpath, false, Encoding.UTF8);
            for (int i = 0; i < allLine.Count; i++)
            {
                if (i == endIndex)
                {
                    sw.Write(sb.ToString());
                }
                sw.WriteLine(allLine[i]);
            }
            sw.Close();
            //属性を戻す
            System.IO.File.SetAttributes(xmlpath, attr);
            return true;
        }

        public static int DeleteConfig(string xmlpath, string guid)
        {
            int delCount = 0;

            /* ファイルの有無チェック */
            if (File.Exists(xmlpath) != true)
            {
                return 0;
            }

            //C:\test.txt の属性を取得する
            System.IO.FileAttributes attr = System.IO.File.GetAttributes(xmlpath);

            //読み取り専用属性があるか調べる
            if ((attr & System.IO.FileAttributes.ReadOnly) ==
                System.IO.FileAttributes.ReadOnly)
            {
                //読み取り専用属性を削除する
                System.IO.File.SetAttributes(xmlpath, attr & (~System.IO.FileAttributes.ReadOnly));
            }

            StreamReader sr = new StreamReader(xmlpath);
            List<string> allLine = new List<string>();
            while (sr.EndOfStream != true)
            {
                allLine.Add(sr.ReadLine());
            }
            sr.Close();

            int guidIndex = -1;
            for (int i = 0; i < allLine.Count; i++)
            {
                string str = allLine[i].Trim(' ', '\t');
                if (str.Contains(guid) == true)
                {
                    guidIndex = i;
                    break;
                }
            }
            if (guidIndex == -1)
            {
                return delCount;
            }
            int startIndex = -1;
            for (int i = guidIndex - 1; i >= 0; i--)
            {
                string str = allLine[i].Trim(' ', '\t');
                if (str.Contains("<AddIn>") == true)
                {
                    startIndex = i;
                    break;
                }
            }
            if (startIndex == -1)
            {
                return delCount;
            }
            int endIndex = -1;
            for (int i = startIndex; i < allLine.Count; i++)
            {
                string str = allLine[i].Trim(' ', '\t');
                if (str.Contains("</AddIn>") == true)
                {
                    endIndex = i;
                    break;
                }
            }
            if (endIndex == -1)
            {
                return delCount;
            }
            delCount = 1;
            StreamWriter sw = new StreamWriter(xmlpath, false, Encoding.UTF8);
            for (int i = 0; i < allLine.Count; i++)
            {
                if ((i >= startIndex) && (i <= endIndex))
                {
                    continue;
                }
                sw.WriteLine(allLine[i]);
            }
            sw.Close();
            //属性を戻す
            System.IO.File.SetAttributes(xmlpath, attr);
            return delCount;
        }
    }

}
