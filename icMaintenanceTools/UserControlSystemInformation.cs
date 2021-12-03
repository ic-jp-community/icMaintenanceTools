using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Management;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ICApiAddin.icMaintenanceTools
{
    public partial class UserControlSystemInformation: UserControl
    {
        public const string title = "システム情報の取得";

        /// <summary>
        /// インストールされているデバイスドライバの情報
        /// </summary>
        private class DEVICE_DRIVER_INFO
        {
            public string manufacturer;
            public string deviceName;
            public string driverVersion;

            public DEVICE_DRIVER_INFO()
            {
                this.manufacturer = string.Empty;
                this.deviceName = string.Empty;
                this.driverVersion = string.Empty;
            }
            public DEVICE_DRIVER_INFO(string manufacturer, string deviceName, string driverVersion)
            {
                this.manufacturer = manufacturer;
                this.deviceName = deviceName;
                this.driverVersion = driverVersion;
            }
        }


        /// <summary>
        /// インストールされているアプリケーションの情報
        /// </summary>
        private class INSTALL_APPLICATION_INFO
        {
            public string publisher;
            public string displayName;
            public string displayVersion;

            public INSTALL_APPLICATION_INFO()
            {
                this.publisher = string.Empty;
                this.displayName = string.Empty;
                this.displayVersion = string.Empty;
            }
            public INSTALL_APPLICATION_INFO(string publisher, string displayName, string displayVersion)
            {
                this.publisher = publisher;
                this.displayName = displayName;
                this.displayVersion = displayVersion;
            }
        }


        /// <summary>
        /// コンストラクタ
        /// </summary>
        public UserControlSystemInformation()
        {
            InitializeComponent();
            this.Tag = new UserControlTagData();
            this.Dock = DockStyle.Fill;
        }


        /// <summary>
        /// ページ表示状態変更イベント(ページ表示イベントとして利用)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void UserControlSystemInformation_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible && !Disposing)
            {
                icapiCommon.enableDataGridViewDoubleBuffer(dataGridViewEventLog);
                ((UserControlTagData)this.Tag).canNotClose = true;
                this.Cursor = Cursors.WaitCursor;
                this.textBoxSystemInformation.Cursor = this.Cursor;
                textBoxSystemInformation.Text = "システム情報を取得中です。しばらくお待ちください。";
                dataGridViewEventLog.Columns.Add("Initialize", "イベント取得中");
                dataGridViewEventLog.Rows.Add(new object[] { "イベント情報を取得中です。しばらくお待ちください。" });
                dataGridViewEventLog.ClearSelection();
                string systemInfo = await ShowSystemInformation();
                textBoxSystemInformation.Text = systemInfo;

                DataTable dt = await getEventLog();
                dataGridViewEventLog.Rows.Clear();
                dataGridViewEventLog.Columns.Clear();
                dataGridViewEventLog.DataSource = dt;
                dataGridViewEventLog.Columns["EventType"].HeaderText = "イベント種別";
                dataGridViewEventLog.Columns["EventTime"].HeaderText = "日時";
                dataGridViewEventLog.Columns["EventSource"].HeaderText = "ソース";
                dataGridViewEventLog.Columns["EventID"].HeaderText = "イベントID";
                dataGridViewEventLog.Columns["EventLog"].HeaderText = "内容";
                this.Cursor = Cursors.Default;
                this.textBoxSystemInformation.Cursor = this.Cursor;
                ((UserControlTagData)this.Tag).canNotClose = false;
            }
        }


        /// <summary>
        /// オペレーティングシステムの情報を取得する
        /// </summary>
        /// <param name="sb"></param>
        private static void getOperatingSystemInformation(ref StringBuilder sb)
        {
            sb.AppendLine("================================================== ");
            sb.AppendLine("      インストールされているWindowsの情報          ");
            sb.AppendLine("================================================== ");
            // Win32_OperatingSystemクラスを作成する
            using (ManagementClass managementClass = new ManagementClass("Win32_OperatingSystem"))
            {
                // Win32_OperatingSystemオブジェクトを取得する
                managementClass.Get();
                // 権限を有効化する
                managementClass.Scope.Options.EnablePrivileges = true;

                // WMIのオブジェクトのコレクションを取得する
                using (ManagementObjectCollection managementObjectCollection = managementClass.GetInstances())
                {
                    // WMIのオブジェクトを列挙する
                    foreach (ManagementObject managementObject in managementObjectCollection)
                    {
                        // OSの名前
                        sb.AppendLine($"Name: {managementObject["Name"]}");
                        // OSの簡単な説明
                        sb.AppendLine($"Caption: {managementObject["Caption"]}");
                        // OSの説明（コンピュータの説明）
                        sb.AppendLine($"Description: {managementObject["Description"]}");
                        // OSのバージョン
                        sb.AppendLine($"Version: {managementObject["Version"]}");
                        // OSのビルド番号
                        sb.AppendLine($"BuildNumber: {managementObject["BuildNumber"]}");
                        // OSの製造元の名前
                        sb.AppendLine($"Manufacturer: {managementObject["Manufacturer"]}");
                        // OSの言語識別子（言語ID）
                        sb.AppendLine($"Locale: {managementObject["Locale"]}");
                        // OSの言語
                        sb.AppendLine($"OSLanguage: {managementObject["OSLanguage"]}");
                        // シリアルナンバー
                        sb.AppendLine($"SerialNumber: {managementObject["SerialNumber"]}");
                        // 製品の種類（1: ワークステーション、2: ドメインコントローラー, 3: サーバー）
                        sb.AppendLine($"ProductType: {managementObject["ProductType"]}");
                        // OSのアーキテクチャ（32ビット or 64 ビット）
                        sb.AppendLine($"OSArchitecture: {managementObject["ProductType"]}");
                        // OSがインストールされた日時
                        sb.AppendLine($"InstallDate: {managementObject["InstallDate"]}");
                        // OSが最後に起動された日時
                        sb.AppendLine($"LastBootUpTime: {managementObject["LastBootUpTime"]}");
                        // OSがインストールされているデバイス（物理ディスク／パーティション）
                        sb.AppendLine($"SystemDevice: {managementObject["SystemDevice"]}");
                        // OSが起動するデバイス（ディスクドライブ）
                        sb.AppendLine($"BootDevice: {managementObject["BootDevice"]}");
                        // システムドライブ
                        sb.AppendLine($"SystemDrive: {managementObject["SystemDrive"]}");
                        // Windowsディレクトリ
                        sb.AppendLine($"WindowsDirectory: {managementObject["WindowsDirectory"]}");
                        //ログイン名
                        sb.AppendLine($"LoginName:" + Environment.UserName);
                        // WMIのオブジェクトのリソースを開放する
                        managementObject.Dispose();
                        break;
                    }
                }
            }
        }


        /// <summary>
        /// グラフィックボードの情報を取得する
        /// </summary>
        /// <param name="sb"></param>
        private static void getGraphicBoardInformation(ref StringBuilder sb)
        {
            sb.AppendLine("================================================== ");
            sb.AppendLine("      グラフィックカードの情報                     ");
            sb.AppendLine("================================================== ");
            using (var searcher = new ManagementObjectSearcher("select * from Win32_VideoController"))
            {
                int boardCount = 0;
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine(string.Format("[BoardNumber: {0}]", boardCount));
                    sb.AppendLine("Name  -  " + obj["Name"]);
                    sb.AppendLine("DeviceID  -  " + obj["DeviceID"]);
                    sb.AppendLine("AdapterRAM  -  " + obj["AdapterRAM"]);
                    sb.AppendLine("AdapterDACType  -  " + obj["AdapterDACType"]);
                    sb.AppendLine("Monochrome  -  " + obj["Monochrome"]);
                    sb.AppendLine("InstalledDisplayDrivers  -  " + obj["InstalledDisplayDrivers"]);
                    sb.AppendLine("DriverVersion  -  " + obj["DriverVersion"]);
                    sb.AppendLine("VideoProcessor  -  " + obj["VideoProcessor"]);
                    sb.AppendLine("VideoArchitecture  -  " + obj["VideoArchitecture"]);
                    sb.AppendLine("VideoMemoryType  -  " + obj["VideoMemoryType"]);
                }
            }
        }


        /// <summary>
        /// 端末のハードウェア情報を取得する
        /// </summary>
        /// <param name="sb"></param>
        private void getMachineHardwareInformation(ref StringBuilder sb)
        {
            sb.AppendLine("================================================== ");
            sb.AppendLine("     端末の情報                                    ");
            sb.AppendLine("================================================== ");
            GetProcessorInformation(ref sb);
            GetMotherBoardInformation(ref sb);
            GetPhysicalMemory(ref sb);
            GetComputerName(ref sb);
        }


        /// <summary>
        /// プロセッサのIDを取得する
        /// </summary>
        /// <param name="sb"></param>
        private void GetProcessorInformation(ref StringBuilder sb)
        {
            ManagementClass mancl = new ManagementClass("win32_processor");
            ManagementObjectCollection manageobcl = mancl.GetInstances();
            double? GHz = null;
            String Id = String.Empty;
            string robotCpu = string.Empty;
            String name = String.Empty;
            foreach (ManagementObject myObj in manageobcl)
            {
                GHz = 0.001 * (UInt32)myObj.Properties["CurrentClockSpeed"].Value;
                Id = myObj.Properties["processorID"].Value.ToString();
                robotCpu = myObj.Properties["Manufacturer"].Value.ToString();
                name = (string)myObj["Name"];
                name = name.Replace("(TM)", "™").Replace("(tm)", "™").Replace("(R)", "®").Replace("(r)", "®").Replace("(C)", "©").Replace("(c)", "©").Replace("    ", " ").Replace("  ", " ");
                sb.AppendLine(string.Format("ProcessorID:{0}", Id));
                sb.AppendLine(string.Format("Manufacturer:{0}", robotCpu));
                sb.AppendLine(string.Format("ClockFreq(GHz):{0}", GHz.ToString()));
                sb.AppendLine(name + ", " + (string)myObj["Caption"] + ", " + (string)myObj["SocketDesignation"]);
                break;
            }
            return;

        }
        //public static String GetHDDSerialNo()
        //{
        //    ManagementClass mangnmt = new ManagementClass("Win32_LogicalDisk");
        //    ManagementObjectCollection mcol = mangnmt.GetInstances();
        //    string result = "";
        //    foreach (ManagementObject strt in mcol)
        //    {
        //        result += Convert.ToString(strt["VolumeSerialNumber"]);
        //    }
        //    return result;
        //}
        //public static string GetMACAddress()
        //{
        //    ManagementClass mancl = new ManagementClass("Win32_NetworkAdapterConfiguration");
        //    ManagementObjectCollection manageobcl = mancl.GetInstances();
        //    string MACAddress = String.Empty;
        //    foreach (ManagementObject myObj in manageobcl)
        //    {
        //        if (MACAddress == String.Empty)
        //        {
        //            if ((bool)myObj["IPEnabled"] == true) MACAddress = myObj["MacAddress"].ToString();
        //        }
        //        myObj.Dispose();
        //    }

        //    MACAddress = MACAddress.Replace(":", "");
        //    return MACAddress;
        //}


        /// <summary>
        /// マザーボードの情報を取得する
        /// </summary>
        /// <param name="sb"></param>
        public void GetMotherBoardInformation(ref StringBuilder sb)
        {
            ManagementObjectSearcher finder = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
            foreach (ManagementObject wmi in finder.Get())
            {
                try
                {
                    sb.AppendLine(string.Format("Manufacturer: {0}", wmi.GetPropertyValue("Manufacturer").ToString()));
                    sb.AppendLine(string.Format("Product: {0}", wmi.GetPropertyValue("Product").ToString()));
                    break;
                }
                catch (Exception ex)
                {
                }
            }
            return;
        }


        /// <summary>
        /// アカウント名を取得する
        /// </summary>
        /// <param name="sb"></param>
        public void GetAccountName(ref StringBuilder sb)
        {
            ManagementObjectSearcher finder = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_UserAccount");
            foreach (ManagementObject wmi in finder.Get())
            {
                try
                {
                    sb.AppendLine(string.Format("UserName: {0}", wmi.GetPropertyValue("Name").ToString()));
                    return;
                }
                catch (Exception ex)
                {
                }
            }
            return;
        }


        /// <summary>
        /// 物理メモリのサイズを取得する
        /// </summary>
        /// <param name="sb"></param>
        public void GetPhysicalMemory(ref StringBuilder sb)
        {
            ManagementScope osmanagesys = new ManagementScope();
            ObjectQuery myobjQry = new ObjectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
            ManagementObjectSearcher oFinder = new ManagementObjectSearcher(osmanagesys, myobjQry);
            ManagementObjectCollection oCollection = oFinder.Get();

            long liveMSize = 0;
            long livecap = 0;

            foreach (ManagementObject obj in oCollection)
            {
                livecap = Convert.ToInt64(obj["Capacity"]);
                liveMSize += livecap;
            }
            liveMSize = (liveMSize / 1024) / 1024;
            sb.AppendLine(string.Format("MemorySize: {0}", liveMSize.ToString() + "MB"));
            return;
        }


        /// <summary>
        /// マシン名を取得する
        /// </summary>
        /// <param name="sb"></param>
        public void GetComputerName(ref StringBuilder sb)
        {
            ManagementClass mancl = new ManagementClass("Win32_ComputerSystem");
            ManagementObjectCollection manageobcl = mancl.GetInstances();
            String info = String.Empty;
            foreach (ManagementObject myObj in manageobcl)
            {
                sb.AppendLine(string.Format("MachineName: {0}", myObj["Name"]));
                break;
            }
            return;
        }

        /// <summary>
        /// デバイスドライバとインストールソフトウェアの情報を取得する
        /// </summary>
        /// <returns></returns>
        private Task<string> ShowSystemInformation()
        {
            return Task.Run<string>(() =>
            {
                StringBuilder sb = new StringBuilder();
                getOperatingSystemInformation(ref sb);
                sb.AppendLine();
                getGraphicBoardInformation(ref sb);
                sb.AppendLine();
                getMachineHardwareInformation(ref sb);
                sb.AppendLine();
                getInstalledDeviceDrivers(ref sb);
                sb.AppendLine();
                getInstalledApplications(ref sb);
                return sb.ToString();
            });
        }

        /// <summary>
        /// インストールされているデバイスドライバの情報を取得する
        /// </summary>
        /// <returns></returns>
        private static void getInstalledDeviceDrivers(ref StringBuilder sb)
        {
            ManagementObjectSearcher objSearcher = new ManagementObjectSearcher("Select * from Win32_PnPSignedDriver");

            ManagementObjectCollection objCollection = objSearcher.Get();
            List<DEVICE_DRIVER_INFO> driverInfoList = new List<DEVICE_DRIVER_INFO>();
            sb.AppendLine("================================================== ");
            sb.AppendLine("     デバイスドライバの情報                        ");
            sb.AppendLine("================================================== ");

            foreach (ManagementObject obj in objCollection)
            {
                string deviceName = string.Empty;
                string Manufacturer = string.Empty;
                string driverVersion = string.Empty;
                if (obj["DeviceName"] != null)
                {
                    deviceName = obj["DeviceName"].ToString();
                }
                if (obj["Manufacturer"] != null)
                {
                    Manufacturer = obj["Manufacturer"].ToString();
                }
                if (obj["DriverVersion"] != null)
                {
                    driverVersion = obj["DriverVersion"].ToString();
                }
                driverInfoList.Add(new DEVICE_DRIVER_INFO(Manufacturer, deviceName, driverVersion));
            }

            List<DEVICE_DRIVER_INFO> sortedDriverInfoList = driverInfoList.OrderBy(a => a.manufacturer).ToList();
            string lastManufacturer = "dummyManufacturer";
            for (int i = 0; i < sortedDriverInfoList.Count; i++)
            {
                string deviceName = sortedDriverInfoList[i].deviceName;
                string manufacturer = sortedDriverInfoList[i].manufacturer;
                string driverVersion = sortedDriverInfoList[i].driverVersion;
                if (string.Equals(lastManufacturer, manufacturer) != true)
                {
                    sb.AppendLine("----------------------------------------");
                    sb.AppendLine(string.Format(" Manufacturer: {0}", manufacturer));
                }
                sb.AppendLine(string.Format("   DeviceName: {0}  (Version: {1})", deviceName, driverVersion));
                lastManufacturer = manufacturer;
            }
            return;
        }

        /// <summary>
        /// インストールされているソフトウェアの情報を取得する
        /// </summary>
        /// <returns></returns>
        private static void getInstalledApplications(ref StringBuilder sb)
        {
            List<INSTALL_APPLICATION_INFO> installAppInfoList = new List<INSTALL_APPLICATION_INFO>();
            sb.AppendLine("================================================== ");
            sb.AppendLine("     インストールされているソフトウェアの情報      ");
            sb.AppendLine("================================================== ");
            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        string publisher = string.Empty;
                        string displayName = string.Empty;
                        string displayVersion = string.Empty;
                        try
                        {
                            publisher = subkey.GetValue("Publisher").ToString();
                        }
                        catch (Exception ex)
                        {
                        }

                        try
                        {
                            displayName = subkey.GetValue("DisplayName").ToString();
                        }
                        catch (Exception ex)
                        {
                        }

                        try
                        {
                            displayVersion = subkey.GetValue("DisplayVersion").ToString();
                        }
                        catch (Exception ex)
                        {
                        }

                        if ((string.IsNullOrEmpty(publisher) == true) &&
                            (string.IsNullOrEmpty(displayName) == true) &&
                            (string.IsNullOrEmpty(displayVersion) == true))
                        {
                            continue;
                        }
                        installAppInfoList.Add(new INSTALL_APPLICATION_INFO(publisher, displayName, displayVersion));
                    }
                }
            }

            List<INSTALL_APPLICATION_INFO> sortedInstallAppInfoList = installAppInfoList.OrderBy(a => a.publisher).ThenBy(a => a.displayName).ThenBy(a => a.displayVersion).ToList();
            string lastPublisher = "dummytPublisher";
            for (int i = 0; i < sortedInstallAppInfoList.Count; i++)
            {
                string publisher = sortedInstallAppInfoList[i].publisher;
                string displayName = sortedInstallAppInfoList[i].displayName;
                string displayVersion = sortedInstallAppInfoList[i].displayVersion;
                if (string.Equals(lastPublisher, publisher) != true)
                {
                    sb.AppendLine("----------------------------------------");
                    sb.AppendLine(string.Format(" Publisher: {0}", publisher));
                }
                sb.AppendLine(string.Format("   Software: {0}  (Version: {1})", displayName, displayVersion));
                lastPublisher = publisher;
            }
            return;
        }

        /// <summary>
        /// イベントログを取得する
        /// </summary>
        /// <returns></returns>
        private async Task<DataTable> getEventLog()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("EventType");
            dt.Columns.Add("EventTime");
            dt.Columns.Add("EventSource");
            dt.Columns.Add("EventID");
            dt.Columns.Add("EventLog");

            bool result =  await Task.Run(() =>
            {
                EventLog logInfo = new EventLog();
                logInfo.Log = "Application";
                logInfo.MachineName = ".";  // Local machine
                string eventType = "";  // Icon for the event
                foreach (EventLogEntry entry in logInfo.Entries.Cast<EventLogEntry>().Reverse<EventLogEntry>())
                {
                    switch (entry.EntryType)
                    {
                        case EventLogEntryType.Warning:
                            eventType = "[Warning]";
                            break;
                        case EventLogEntryType.Error:
                            eventType = "[Error]";
                            break;
                        default:
                            eventType = "[Info]";
                            break;
                    }
                    DataRow dr = dt.NewRow();
                    dr["EventType"] = eventType;
                    dr["EventTime"] = entry.TimeGenerated.ToString();
                    dr["EventSource"] = entry.Source;
                    dr["EventID"] = (UInt16)(entry.InstanceId & 0x3FFFFFFF);
                    dr["EventLog"] = entry.Message;
                    dt.Rows.Add(dr);
                }
                return true;
            });
            return dt;
        }

    }
}
