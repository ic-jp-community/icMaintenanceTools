using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ICApiAddin.icMaintenanceTools
{
    /// <summary>
    /// ユーザーコントロールのTag格納データ
    /// </summary>
    public class UserControlTagData
    {
        public bool canNotClose;
        public UserControlTagData()
        {
            this.canNotClose = false;
        }
    }

    /// <summary>
    /// IRONCADの各フォルダ・レジストリパス
    /// </summary>
    public class IRONCAD_REG_AND_DIR
    {
        public string version { get; set; }
        public string installDir { get; set; }
        public string cacheDirRoaming;
        public string cacheDirLocal;
        public string cacheDirRoamingIRONCAD;
        public string cacheDirLocalIRONCAD;
        public string cacheDirRoamingINOVATE;
        public string cacheDirLocalINOVATE;
        public string cacheDirRoamingCOMPOSE;
        public string cacheDirLocalCOMPOSE;
        public string cacheDirRoamingDRAFT;
        public string cacheDirLocalDRAFT;

        public string cacheDirRoamingCaxaIC;
        public string cacheDirLocalCaxaIC;
        public string cacheDirRoamingCaxaIntDRAFT;
        public string cacheDirLocalCaxaIntDRAFT;
        public string cacheDirRoamingCaxaIntINOVATE;
        public string cacheDirLocalCaxaIntINOVATE;
        public string cacheDirRoamingCaxaIntIC;
        public string cacheDirLocalCaxaIntIC;
        public string cacheDirRoamingCaxaPrintTool;
        public string cacheDirLocalCaxaPrintTool;
        public string cacheDirRoamingCaxaExbViewr;
        public string cacheDirLocalCaxaExbViewr;

        public string regPathUserIRONCAD;
        public string regPathMachineIRONCAD;
        public string regPathMachineIRONCAD_NewUser;
        public string regPathUserINOVATE;
        public string regPathMachineINOVATE;
        public string regPathMachineINOVATE_NewUser;
        public string regPathUserCOMPOSE;
        public string regPathMachineCOMPOSE;
        public string regPathMachineCOMPOSE_NewUser;
        public string regPathUserDRAFT;
        public string regPathMachineDRAFT;
        public string regPathMachineDRAFT_NewUser;

        public IRONCAD_REG_AND_DIR()
        {
            this.version = string.Empty;
            this.installDir = string.Empty;
            this.cacheDirRoaming = string.Empty;
            this.cacheDirLocal = string.Empty;
            this.cacheDirRoamingIRONCAD = string.Empty;
            this.cacheDirLocalIRONCAD = string.Empty;
            this.cacheDirRoamingINOVATE = string.Empty;
            this.cacheDirLocalINOVATE = string.Empty;
            this.cacheDirRoamingCOMPOSE = string.Empty;
            this.cacheDirLocalCOMPOSE = string.Empty;
            this.cacheDirRoamingDRAFT = string.Empty;
            this.cacheDirLocalDRAFT = string.Empty;

            this.cacheDirRoamingCaxaIC = string.Empty;
            this.cacheDirRoamingCaxaIntDRAFT = string.Empty;
            this.cacheDirRoamingCaxaIntINOVATE = string.Empty;
            this.cacheDirRoamingCaxaIntIC = string.Empty;
            this.cacheDirRoamingCaxaPrintTool = string.Empty;
            this.cacheDirRoamingCaxaExbViewr = string.Empty;

            this.cacheDirLocalCaxaIC = string.Empty;
            this.cacheDirLocalCaxaIntDRAFT = string.Empty;
            this.cacheDirLocalCaxaIntINOVATE = string.Empty;
            this.cacheDirLocalCaxaIntIC = string.Empty;
            this.cacheDirLocalCaxaPrintTool = string.Empty;
            this.cacheDirLocalCaxaExbViewr = string.Empty;

            this.regPathUserIRONCAD = string.Empty;
            this.regPathMachineIRONCAD = string.Empty;
            this.regPathMachineIRONCAD_NewUser = string.Empty;
            this.regPathUserINOVATE = string.Empty;
            this.regPathMachineINOVATE = string.Empty;
            this.regPathMachineINOVATE_NewUser = string.Empty;
            this.regPathUserCOMPOSE = string.Empty;
            this.regPathMachineCOMPOSE = string.Empty;
            this.regPathMachineCOMPOSE_NewUser = string.Empty;
            this.regPathUserDRAFT = string.Empty;
            this.regPathMachineDRAFT = string.Empty;
            this.regPathMachineDRAFT_NewUser = string.Empty;
        }
    }

    /// <summary>
    /// 共通処理
    /// </summary>
    public static class icapiCommon
    {
        /// <summary>
        /// DataGridViewのダブルバッファを有効にする
        /// </summary>
        /// <param name="dgv"></param>
        public static void enableDataGridViewDoubleBuffer(DataGridView dgv)
        {
            try
            {
                // DataGirdViewのTypeを取得
                System.Type dgvtype = typeof(DataGridView);

                // プロパティ設定の取得
                System.Reflection.PropertyInfo dgvPropertyInfo =
                dgvtype.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

                // 対象のDataGridViewにtrueをセットする
                dgvPropertyInfo.SetValue(dgv, true, null);
            }
            catch (Exception ex)
            {
            }

        }

        /// <summary>
        /// インストールされている全てのバージョンのIRONCADインストールディレクトリを取得する
        /// </summary>
        /// <param name="installedIcList">key:バージョン  Value:インストールパス</param>
        /// <param name="getYearVersion">バージョンを年表記にするか  true:年  false:バージョン</param>
        /// <param name="sortNewVersion">新しいバージョン順にソートする  true:ソートする  false:ソートしない(古い順ソート)</param>
        public static void GetAllIronCADInstallPath(ref List<KeyValuePair<string, string>> installedIcList, bool getYearVersion, bool sortNewVersion)
        {
            installedIcList.Clear();
            List<KeyValuePair<string, string>> tmpInstalledIcList = new List<KeyValuePair<string, string>>();

            /* レジストリから探す */
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\IronCAD\IRONCAD\", false);
            if (rkey != null)
            {
                string[] IronCADKeys = rkey.GetSubKeyNames();
                for (int i = 0; i < IronCADKeys.Count(); i++)
                {
                    Microsoft.Win32.RegistryKey subKey = rkey.OpenSubKey(IronCADKeys[i], false);
                    if (subKey == null)
                    {
                        continue;
                    }
                    string version = string.Empty;
                    if (getYearVersion == true)
                    {
                        version = convertNumberVersionToYearVersion(double.Parse(IronCADKeys[i]));
                    }
                    else
                    {
                        version = IronCADKeys[i];
                    }
                    string installDir = string.Empty;
                    bool installDirValueExists = subKey.GetValueNames().ToList().Contains("InstallDir");
                    if (installDirValueExists == true)
                    {
                        installDir = subKey.GetValue("InstallDir").ToString();
                    }
                    subKey.Close();
                    tmpInstalledIcList.Add(new KeyValuePair<string, string>(version, installDir));
                }
                rkey.Close();
            }

            /* ProgramFilesから探す */
            string programFilesPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string ironcadFolderPath = Path.Combine(programFilesPath, "IronCAD");
            if(Directory.Exists(ironcadFolderPath) == true)
            {
                DirectoryInfo dinfo = new DirectoryInfo(ironcadFolderPath);
                DirectoryInfo[] subDirList = dinfo.GetDirectories();
                foreach(DirectoryInfo subDir in subDirList)
                {
                    int yearVersion = -1;
                    string installDir = subDir.FullName;
                    string dirName = Path.GetFileName(installDir);
                    bool parse = Int32.TryParse(dirName, out yearVersion);
                    if(parse != true)
                    {
                        continue;
                    }
                    string version = string.Empty;
                    if (getYearVersion == true)
                    {
                        version = yearVersion.ToString();
                    }
                    else
                    {
                        version = convertYearVersionToNumberVersion(yearVersion);
                    }
                    int count = tmpInstalledIcList.Where(a => string.Equals(a.Key, version)).Count();
                    if(count != 0)
                    {
                        continue;
                    }
                    tmpInstalledIcList.Add(new KeyValuePair<string, string>(version, installDir));
                }
            }

            /* 結果をソートする */
            if (sortNewVersion == true)
            {
                installedIcList = tmpInstalledIcList.OrderByDescending(pair => pair.Key).ToList();
            }
            else
            {
                installedIcList = tmpInstalledIcList.OrderBy(pair => pair.Key).ToList();
            }
        }


        /// <summary>
        /// IRONCADバージョン(23.0など)を年バージョン(2021など)に変換する
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static string convertNumberVersionToYearVersion(double number)
        {
            int yearVersion = (int)(number + 1998);
            return yearVersion.ToString();
        }


        /// <summary>
        /// IRONCAD年バージョン(2021など)をバージョン(23.0など)に変換する
        /// </summary>
        /// <param name="yearVersion"></param>
        /// <returns></returns>
        public static string convertYearVersionToNumberVersion(int yearVersion)
        {
            double version = (double)(yearVersion - 1998);
            return version.ToString("F1");
        }


        /// <summary>
        /// インストールされている全IRONCADのフォルダ・レジストリ情報を取得する
        /// </summary>
        /// <returns></returns>
        public static List<IRONCAD_REG_AND_DIR> getAllIroncadInformation()
        {
            List<IRONCAD_REG_AND_DIR> ironcadInfomationList = new List<IRONCAD_REG_AND_DIR>();
            string appDataRoamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appDataLocalPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            List<KeyValuePair<string, string>> ironcadList = new List<KeyValuePair<string, string>>();
            icapiCommon.GetAllIronCADInstallPath(ref ironcadList, true, true);
            foreach (KeyValuePair<string, string> ironcad in ironcadList)
            {
                string version = icapiCommon.convertYearVersionToNumberVersion(Int32.Parse(ironcad.Key));
                string yearVersion = ironcad.Key;
                string cacheDirRoaming = appDataRoamingPath;
                string cacheDirLocal = appDataLocalPath;
                string cacheDirRoamingINOVATE = Path.Combine(appDataRoamingPath, @"IronCAD\INOVATE", version);
                string cacheDirLocalINOVATE = Path.Combine(appDataLocalPath, @"IronCAD\INOVATE", version);
                string cacheDirRoamingIRONCAD = Path.Combine(appDataRoamingPath, @"IronCAD\IRONCAD", version);
                string cacheDirLocalIRONCAD = Path.Combine(appDataLocalPath, @"IronCAD\IRONCAD", version);
                string cacheDirRoamingCOMPOSE = Path.Combine(appDataRoamingPath, @"IronCAD\IRONCAD COMPOSE", version);
                string cacheDirLocalCOMPOSE = Path.Combine(appDataLocalPath, @"IronCAD\IRONCAD COMPOSE", version);
                string cacheDirRoamingDRAFT = Path.Combine(appDataRoamingPath, @"IronCAD\IRONCAD DRAFT", version);
                string cacheDirLocalDRAFT = Path.Combine(appDataLocalPath, @"IronCAD\IRONCAD DRAFT", version);

                string cacheDirRoamingCaxaIC = Path.Combine(appDataRoamingPath, @"CAXA\CAXA CAD CXIC " + yearVersion + " (x64)");
                string cacheDirRoamingCaxaIntDRAFT = Path.Combine(appDataRoamingPath, @"CAXA\CAXA CAD CXIC-INT-DRAFT " + yearVersion + " (x64)");
                string cacheDirRoamingCaxaIntINOVATE = Path.Combine(appDataRoamingPath, @"CAXA\CAXA CAD CXIC-INT-INOVATE " + yearVersion + " (x64)");
                string cacheDirRoamingCaxaIntIC = Path.Combine(appDataRoamingPath, @"CAXA\CAXA CAD CXIC-INT-IRONCAD " + yearVersion + " (x64)");
                string cacheDirRoamingCaxaPrintTool = Path.Combine(appDataRoamingPath, @"CAXA\CAXA PRINT TOOL " + yearVersion + " (x64)");
                string cacheDirRoamingCaxaExbViewr = Path.Combine(appDataRoamingPath, @"CAXA\CAXA EXB VIEWER " + yearVersion + " (x64)");

                string cacheDirLocalCaxaIC = Path.Combine(appDataLocalPath, @"CAXA\CAXA CAD CXIC " + yearVersion + " (x64)");
                string cacheDirLocalCaxaIntDRAFT = Path.Combine(appDataLocalPath, @"CAXA\CAXA CAD CXIC-INT-DRAFT " + yearVersion + " (x64)");
                string cacheDirLocalCaxaIntINOVATE = Path.Combine(appDataLocalPath, @"CAXA\CAXA CAD CXIC-INT-INOVATE " + yearVersion + " (x64)");
                string cacheDirLocalCaxaIntIC = Path.Combine(appDataLocalPath, @"CAXA\CAXA CAD CXIC-INT-IRONCAD " + yearVersion + " (x64)");
                string cacheDirLocalCaxaPrintTool = Path.Combine(appDataLocalPath, @"CAXA\CAXA PRINT TOOL " + yearVersion + " (x64)");
                string cacheDirLocalCaxaExbViewr = Path.Combine(appDataLocalPath, @"CAXA\CAXA EXB VIEWER " + yearVersion + " (x64)");

                string regPathUserINOVATE = @"SOFTWARE\IronCAD\INOVATE " + version;
                string regPathMachineINOVATE = @"SOFTWARE\IronCAD\INOVATE\" + version;
                string regPathMachineINOVATE_NewUser = @"SOFTWARE\IronCAD\New User Settings\INOVATE " + version;
                string regPathUserIRONCAD = @"SOFTWARE\IronCAD\IRONCAD " + version;
                string regPathMachineIRONCAD = @"SOFTWARE\IronCAD\IRONCAD\" + version;
                string regPathMachineIRONCAD_NewUser = @"SOFTWARE\IronCAD\New User Settings\IRONCAD " + version;
                string regPathUserCOMPOSE = @"SOFTWARE\IronCAD\IRONCAD COMPOSE " + version;
                string regPathMachineCOMPOSE = @"SOFTWARE\IronCAD\IRONCAD COMPOSE\" + version;
                string regPathMachineCOMPOSE_NewUser = @"SOFTWARE\IronCAD\New User Settings\IRONCAD COMPOSE " + version;
                string regPathUserDRAFT = @"SOFTWARE\IronCAD\IRONCAD DRAFT " + version;
                string regPathMachineDRAFT = @"SOFTWARE\IronCAD\IRONCAD DRAFT\" + version;
                string regPathMachineDRAFT_NewUser = @"SOFTWARE\IronCAD\New User Settings\IRONCAD DRAFT " + version;

                IRONCAD_REG_AND_DIR regAndDir = new IRONCAD_REG_AND_DIR();
                regAndDir.version = ironcad.Key;
                regAndDir.installDir = ironcad.Value;
                regAndDir.cacheDirLocal = cacheDirLocal;
                regAndDir.cacheDirRoaming = cacheDirRoaming;
                regAndDir.cacheDirLocalCOMPOSE = existDirectoryPath(cacheDirLocalCOMPOSE);
                regAndDir.cacheDirRoamingCOMPOSE = existDirectoryPath(cacheDirRoamingCOMPOSE);
                regAndDir.cacheDirLocalIRONCAD = existDirectoryPath(cacheDirLocalIRONCAD);
                regAndDir.cacheDirRoamingIRONCAD = existDirectoryPath(cacheDirRoamingIRONCAD);
                regAndDir.cacheDirLocalINOVATE = existDirectoryPath(cacheDirLocalINOVATE);
                regAndDir.cacheDirRoamingINOVATE = existDirectoryPath(cacheDirRoamingINOVATE);
                regAndDir.cacheDirLocalDRAFT = existDirectoryPath(cacheDirLocalDRAFT);
                regAndDir.cacheDirRoamingDRAFT = existDirectoryPath(cacheDirRoamingDRAFT);

                regAndDir.cacheDirRoamingCaxaIC = existDirectoryPath(cacheDirRoamingCaxaIC);
                regAndDir.cacheDirRoamingCaxaIntDRAFT = existDirectoryPath(cacheDirRoamingCaxaIntDRAFT);
                regAndDir.cacheDirRoamingCaxaIntINOVATE = existDirectoryPath(cacheDirRoamingCaxaIntINOVATE);
                regAndDir.cacheDirRoamingCaxaIntIC = existDirectoryPath(cacheDirRoamingCaxaIntIC);
                regAndDir.cacheDirRoamingCaxaPrintTool = existDirectoryPath(cacheDirRoamingCaxaPrintTool);
                regAndDir.cacheDirRoamingCaxaExbViewr = existDirectoryPath(cacheDirRoamingCaxaExbViewr);

                regAndDir.cacheDirLocalCaxaIC = existDirectoryPath(cacheDirLocalCaxaIC);
                regAndDir.cacheDirLocalCaxaIntDRAFT = existDirectoryPath(cacheDirLocalCaxaIntDRAFT);
                regAndDir.cacheDirLocalCaxaIntINOVATE = existDirectoryPath(cacheDirLocalCaxaIntINOVATE);
                regAndDir.cacheDirLocalCaxaIntIC = existDirectoryPath(cacheDirLocalCaxaIntIC);
                regAndDir.cacheDirLocalCaxaPrintTool = existDirectoryPath(cacheDirLocalCaxaPrintTool);
                regAndDir.cacheDirLocalCaxaExbViewr = existDirectoryPath(cacheDirLocalCaxaExbViewr);

                regAndDir.regPathUserCOMPOSE = existsRegistoryPath(regPathUserCOMPOSE, false);
                regAndDir.regPathMachineCOMPOSE = existsRegistoryPath(regPathMachineCOMPOSE, true);
                regAndDir.regPathMachineCOMPOSE_NewUser = existsRegistoryPath(regPathMachineCOMPOSE_NewUser, true);
                regAndDir.regPathUserIRONCAD = existsRegistoryPath(regPathUserIRONCAD, false);
                regAndDir.regPathMachineIRONCAD = existsRegistoryPath(regPathMachineIRONCAD, true);
                regAndDir.regPathMachineIRONCAD_NewUser = existsRegistoryPath(regPathMachineIRONCAD_NewUser, true);
                regAndDir.regPathUserINOVATE = existsRegistoryPath(regPathUserINOVATE, false);
                regAndDir.regPathMachineINOVATE = existsRegistoryPath(regPathMachineINOVATE, true);
                regAndDir.regPathMachineINOVATE_NewUser = existsRegistoryPath(regPathMachineINOVATE_NewUser, true);
                regAndDir.regPathUserDRAFT = existsRegistoryPath(regPathUserDRAFT, false);
                regAndDir.regPathMachineDRAFT = existsRegistoryPath(regPathMachineDRAFT, true);
                regAndDir.regPathMachineDRAFT_NewUser = existsRegistoryPath(regPathMachineDRAFT_NewUser, true);
                ironcadInfomationList.Add(regAndDir);
            }
            return ironcadInfomationList;
        }


        /// <summary>
        /// レジストリがあるかチェックする
        /// </summary>
        /// <param name="regPath">チェックするレジストリパス</param>
        /// <param name="isMachine">レジストリ種別  true:端末(MACHINE)  false: ユーザ(CURRENT_USER)</param>
        /// <returns>空白: レジストリなし  入力パスと同じ: レジストリあり</returns>
        public static string existsRegistoryPath(string regPath, bool isMachine)
        {
            if (isMachine == true)
            {
                Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regPath, false);
                if (rkey == null)
                {
                    return string.Empty;
                }
                rkey.Close();
            }
            else
            {
                Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regPath, false);
                if (rkey == null)
                {
                    return string.Empty;
                }
                rkey.Close();
            }
            return regPath;
        }


        /// <summary>
        /// ディレクトリがあるかチェックする
        /// </summary>
        /// <param name="dirPath">チェックするディレクトリパス</param>
        /// <returns>空白: ディレクトリなし  入力パスと同じ: ディレクトリあり</returns>
        public static string existDirectoryPath(string dirPath)
        {
            if (Directory.Exists(dirPath) != true)
            {
                return string.Empty;
            }
            return dirPath;
        }


        /// <summary>
        /// 読み取り専用を解除する(配下のファイルディレクトリすべて)
        /// </summary>
        /// <param name="dirInfo">解除するディレクトリトップのパス</param>
        public static void RemoveReadonlyAttribute(DirectoryInfo dirInfo)
        {
            //基のフォルダの属性を変更
            if ((dirInfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                dirInfo.Attributes = FileAttributes.Normal;
            }

            //フォルダ内のすべてのファイルの属性を変更
            foreach (FileInfo fi in dirInfo.GetFiles())
            {
                if ((fi.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    fi.Attributes = FileAttributes.Normal;
            }

            //サブフォルダの属性を回帰的に変更
            foreach (DirectoryInfo di in dirInfo.GetDirectories())
            {
                RemoveReadonlyAttribute(di);
            }
        }


        /// <summary>
        /// ディレクトリを削除する(配下のディレクトリすべて)
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="maxRetries"></param>
        /// <param name="millisecondsDelay"></param>
        /// <returns></returns>
        public static async Task<bool> deleteDirectory(string directoryPath, int maxRetries = 10, int millisecondsDelay = 30)
        {
            if ((string.IsNullOrEmpty(directoryPath) == true) ||
                (Directory.Exists(directoryPath) != true))
            {
                return true;
            }

            /* 読み取り専用を解除する */
            RemoveReadonlyAttribute(new DirectoryInfo(directoryPath));

            bool ret = false;
            for (int i = 0; i < maxRetries; ++i)
            {
                try
                {
                    if (Directory.Exists(directoryPath))
                    {
                        Directory.Delete(directoryPath, true);
                    }
                    ret = true;
                    break;
                }
                catch (IOException ex)
                {
                    await Task.Delay(millisecondsDelay);
                }
                catch (UnauthorizedAccessException ex)
                {
                    await Task.Delay(millisecondsDelay);
                }
            }

            return ret;
        }


        /// <summary>
        /// ディレクトリをバックアップ(コピー)する
        /// </summary>
        /// <param name="directoryPath">コピー元</param>
        /// <param name="backupPath">コピー先</param>
        /// <returns></returns>
        public static async Task<bool> BackupDirectory(string directoryPath, string backupPath)
        {
            return await Task.Run<bool>(() =>
            {
                bool result = false;
                if ((string.IsNullOrEmpty(directoryPath) == true) ||
                   (Directory.Exists(directoryPath) != true))
                {
                    return true;
                }
                try
                {
                    Copy(directoryPath, backupPath);
                    result = true;
                }
                catch (Exception ex)
                {

                }
                return result;
            });
        }

        private static void Copy(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        private static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            foreach (FileInfo fi in source.GetFiles())
            {
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
    }
}
