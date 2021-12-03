using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ICApiAddin.icMaintenanceTools
{
    public partial class UserControlIroncadRepair : UserControl
    {
        public const string title = "特定のIRONCADバージョンの修復";
        public List<IRONCAD_REG_AND_DIR> _ironcadList = new List<IRONCAD_REG_AND_DIR>();
        List<string> _workCacheRoamingPaths = new List<string>();
        List<string> _workCacheLocalPaths = new List<string>();
        List<string> _workUserRegPaths = new List<string>();
        List<string> _uninstallCacheLocalPaths = new List<string>();
        List<string> _uninstallCacheRoamingPaths = new List<string>();
        List<string> _uninstallUserRegPaths = new List<string>();
        List<string> _uninstallInstallPaths = new List<string>();
        List<string> _uninstallMachineRegPaths = new List<string>();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public UserControlIroncadRepair()
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
        private void UserControlIroncadRepair_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible && !Disposing)
            {
                this.Cursor = Cursors.WaitCursor;
                initializeControlAndInnerData(string.Empty);
                this.Cursor = Cursors.Default;
            }
        }


        /// <summary>
        /// コントロールと内部データを初期化する
        /// </summary>
        /// <param name="comboBoxSelectVersion">選択するバージョンを指定(空白の場合は初期選択値)</param>
        private void initializeControlAndInnerData(string comboBoxSelectVersion)
        {
            this._ironcadList.Clear();
            this._ironcadList = icapiCommon.getAllIroncadInformation();

            comboBoxIroncadVersion.DataSource = this._ironcadList;
            comboBoxIroncadVersion.DisplayMember = "version";
            comboBoxIroncadVersion.ValueMember = "installDir";

            if(string.IsNullOrEmpty(comboBoxSelectVersion) != true) 
            {
                IEnumerable<IRONCAD_REG_AND_DIR> found = this._ironcadList.Where(a => string.Equals(a.version, comboBoxSelectVersion));
                if(found != null)
                {
                    comboBoxIroncadVersion.Text = comboBoxSelectVersion;
                }
            }
            refreshDeletePathAndRegList();
        }


        /// <summary>
        /// ユーザーのキャッシュデータをバックアップし削除する
        /// </summary>
        /// <param name="ironcadPaths"></param>
        /// <param name="roamingPaths">削除するAppData\Roamingパス</param>
        /// <param name="localPaths">削除するAppData\Localパス</param>
        /// <param name="backupTopPath"></param>
        /// <returns></returns>
        private async Task<bool> deleteAndBackupUserCache(IRONCAD_REG_AND_DIR ironcadPaths, List<string> roamingPaths, List<string> localPaths, string backupTopPath)
        {
            /* Roaming */
            foreach (string cachePath in roamingPaths)
            {
                if ((string.IsNullOrEmpty(cachePath) == false &&
                    (Directory.Exists(cachePath) == true)))
                {
                    string r_baseFolderName = Path.GetFileName(ironcadPaths.cacheDirRoaming); /* Roaming */
                    string r_prgName = cachePath.Replace(ironcadPaths.cacheDirRoaming, string.Empty).Trim('\\'); /* IRONCAD */
                    string r_backupPath = Path.Combine(backupTopPath, r_baseFolderName, r_prgName);

                    bool r_result = await icapiCommon.BackupDirectory(cachePath, r_backupPath);
                    if (r_result != true)
                    {
                        return false;
                    }
                    r_result = await icapiCommon.deleteDirectory(cachePath);
                    if (r_result != true)
                    {
                        return false;
                    }
                }
            }

            /* Local */
            foreach(string cachePath in localPaths)
            {
                if ((string.IsNullOrEmpty(cachePath) == false) &&
                    (Directory.Exists(cachePath) == true))
                {
                    string l_baseFolderName = Path.GetFileName(ironcadPaths.cacheDirLocal); /* Local */
                    string l_prgName = cachePath.Replace(ironcadPaths.cacheDirLocal, string.Empty).Trim('\\'); /* IRONCAD */
                    string l_backupPath = Path.Combine(backupTopPath, l_baseFolderName, l_prgName);

                    bool l_result = await icapiCommon.BackupDirectory(cachePath, l_backupPath);
                    if (l_result != true)
                    {
                        return false;
                    }
                    l_result = await icapiCommon.deleteDirectory(cachePath);
                    if (l_result != true)
                    {
                        return false;
                    }
                }
            }
            return true;
        }


        /// <summary>
        /// レジストリをエクスポートする
        /// </summary>
        /// <param name="isMachine">レジストリ種別  true:端末(MACHINE)  false: ユーザ(CURRENT_USER)</param>
        /// <param name="strKey">レジストリパス</param>
        /// <param name="exportFilePath">バックアップファイルパス</param>
        /// <returns></returns>
        private async Task<bool> exportRegistry(bool isMachine, string strKey, string exportFilePath)
        {
            bool ret =  await Task.Run<bool>(() =>
            {
                bool result = false;
                try
                {
                    string regType = string.Empty;
                    if (isMachine == true)
                    {
                        regType = "HKEY_LOCAL_MACHINE";
                    }
                    else
                    {
                        regType = "HKEY_CURRENT_USER";
                    }
                    strKey = Path.Combine(regType, strKey);
                    using (Process proc = new Process())
                    {
                        proc.StartInfo.FileName = "reg.exe";
                        proc.StartInfo.UseShellExecute = false;
                        proc.StartInfo.RedirectStandardOutput = true;
                        proc.StartInfo.RedirectStandardError = true;
                        proc.StartInfo.CreateNoWindow = true;
                        proc.StartInfo.Arguments = "export \"" + strKey + "\" \"" + exportFilePath + "\" /y";
                        proc.Start();
                        string stdout = proc.StandardOutput.ReadToEnd();
                        string stderr = proc.StandardError.ReadToEnd();
                        proc.WaitForExit();
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    // handle exception
                    result = false;
                }
                return result;
            });
            return ret;
        }


        /// <summary>
        /// レジストリを削除する
        /// </summary>
        /// <param name="isMachine">レジストリ種別  true:端末(MACHINE)  false: ユーザ(CURRENT_USER)</param>
        /// <param name="strKey">レジストリパス</param>
        /// <returns></returns>
        private async Task<bool> deleteRegistry(bool isMachine, string strKey)
        {
            bool ret = await Task.Run<bool>(() =>
            {
                bool result = false;
                try
                {
                    string existRegPath = icapiCommon.existsRegistoryPath(strKey, isMachine);
                    if(string.IsNullOrEmpty(existRegPath) == true)
                    {
                        return true;
                    }

                    string parent = Path.GetDirectoryName(strKey);
                    string subKey = Path.GetFileName(strKey);
                    if (isMachine == true)
                    {
                        Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(parent, true);
                        if (rkey == null)
                        {
                            return true;
                        }
                        rkey.DeleteSubKeyTree(subKey, false);
                        rkey.Close();
                    }
                    else
                    {
                        Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(parent, true);
                        if (rkey == null)
                        {
                            return true;
                        }
                        rkey.DeleteSubKeyTree(subKey, false);
                        rkey.Close();
                    }
                    result = true;
                }
                catch (Exception ex)
                {
                    // handle exception
                    result = false;
                }
                return result;
            });
            return ret;
        }


        /// <summary>
        /// ユーザのレジストリをエクスポートして削除する
        /// </summary>
        /// <param name="ironcadPaths"></param>
        /// <param name="userRegs">削除対象のユーザレジストリ</param>
        /// <param name="backupTopPath">レジストリのバックアップディレクトリ</param>
        /// <returns></returns>
        private async Task<bool> deleteAndExportUserRegistory(IRONCAD_REG_AND_DIR ironcadPaths, List<string> userRegs, string backupTopPath)
        {
            bool result = true;
            foreach (string regKey in userRegs)
            {
                string baseFolderName = "Registory"; /* Roaming */
                string prgName = "HKEY_CURRENT_USER\\" + Path.GetFileName(regKey) + ".reg";
                string backupPath = Path.Combine(backupTopPath, baseFolderName, prgName);
                DirectoryInfo diTarget = new DirectoryInfo(backupPath);
                if (Directory.Exists(Path.GetDirectoryName(diTarget.FullName)) != true)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(diTarget.FullName));
                }

                result = await exportRegistry(false, regKey, diTarget.FullName);
                if (result != true)
                {
                    result = false;
                    break;
                }
                result = await deleteRegistry(false, regKey);
                if (result != true)
                {
                    result = false;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// 端末のレジストリをエクスポートして削除する
        /// </summary>
        /// <param name="ironcadPaths"></param>
        /// <param name="backupTopPath"></param>
        /// <returns></returns>
        private async Task<bool> deleteAndExportMachineRegistory(IRONCAD_REG_AND_DIR ironcadPaths, List<string> machineRegs, string backupTopPath)
        {
            bool result = true;
            foreach (string regKey in machineRegs)
            {
                string baseFolderName = "Registory"; /* Roaming */
                string name1 = Path.GetFileName(Path.GetDirectoryName(regKey));
                string name2 = Path.GetFileName(regKey);
                string prgName = "HKEY_LOCAL_MACHINE\\" + name1 + "_" + name2 + ".reg";
                string backupPath = Path.Combine(backupTopPath, baseFolderName, prgName);
                DirectoryInfo diTarget = new DirectoryInfo(backupPath);
                if (Directory.Exists(Path.GetDirectoryName(diTarget.FullName)) != true)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(diTarget.FullName));
                }

                result = await exportRegistry(true, regKey, diTarget.FullName);
                if (result != true)
                {
                    result = false;
                    break;
                }
                result = await deleteRegistry(true, regKey);
                if (result != true)
                {
                    result = false;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// IRONCADのインストールフォルダを削除する
        /// </summary>
        /// <param name="ironcadPaths"></param>
        /// <param name="uninstallPaths"></param>
        /// <returns></returns>
        private async Task<bool> deleteIRONCADFolder(IRONCAD_REG_AND_DIR ironcadPaths, List<string> uninstallPaths)
        {
            bool result = true;
            /* Roaming */
            foreach (string installPath in uninstallPaths)
            {
                if ((string.IsNullOrEmpty(installPath) == false &&
                    (Directory.Exists(installPath) == true)))
                {
                    bool r_result = await icapiCommon.deleteDirectory(installPath);
                    if (r_result != true)
                    {
                        result = false;
                        break;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 選択しているIRONCADバージョンを取得する
        /// </summary>
        /// <returns></returns>
        private IRONCAD_REG_AND_DIR getSelectIRONCADVersion()
        {
            string version = comboBoxIroncadVersion.Text;
            IRONCAD_REG_AND_DIR data = this._ironcadList.Where(a => a.version == version).FirstOrDefault();
            return data;
        }


        /// <summary>
        /// ユーザーデータの削除ボタン クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void buttonWorkDelete_Click(object sender, EventArgs e)
        {
            string documentPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string backupTopPath = Path.Combine(documentPath, "IRONCAD JP Community\\icMaintenanceTools", "backup" + DateTime.Now.ToString("yyyyMMddHHmmss"));
            IRONCAD_REG_AND_DIR ironcadPaths = getSelectIRONCADVersion();
            if(ironcadPaths == null)
            {
                return;
            }
            ((UserControlTagData)this.Tag).canNotClose = true;
            buttonWorkDelete.Enabled = false;

            /* ユーザーのキャッシュデータを削除する */
            if (checkBoxWorkDeleteCache.Checked == true)
            {
                bool result = await deleteAndBackupUserCache(ironcadPaths, this._workCacheRoamingPaths, this._workCacheLocalPaths, backupTopPath);
                if (result != true)
                {
                    MessageBox.Show("ユーザーキャッシュの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("ユーザーキャッシュを削除しました。");
                }
            }

            /* ユーザーのレジストリデータを削除する */
            if (checkBoxWorkDeleteUserRegistory.Checked == true)
            {
                bool result = await deleteAndExportUserRegistory(ironcadPaths, this._workUserRegPaths, backupTopPath);
                if (result != true)
                {
                    MessageBox.Show("ユーザーレジストリの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("ユーザーレジストリを削除しました。");
                }
            }
            ((UserControlTagData)this.Tag).canNotClose = false;
            buttonWorkDelete.Enabled = true;
            initializeControlAndInnerData(comboBoxIroncadVersion.Text);
        }


        /// <summary>
        /// アンインストール後のデータ削除 実行ボタンクリック イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void buttonUninstallDelete_Click(object sender, EventArgs e)
        {
            string documentPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string backupTopPath = Path.Combine(documentPath, "IRONCAD JP Community\\icMaintenanceTools", "backup" + DateTime.Now.ToString("yyyyMMddHHmmss"));
            IRONCAD_REG_AND_DIR ironcadPaths = getSelectIRONCADVersion();
            if (ironcadPaths == null)
            {
                return;
            }

            ((UserControlTagData)this.Tag).canNotClose = true;
            buttonUninstallDelete.Enabled = false;

            /* ユーザーのキャッシュデータを削除する */
            if (checkBoxUninstallDeleteCache.Checked == true)
            {
                bool result = await deleteAndBackupUserCache(ironcadPaths, this._uninstallCacheRoamingPaths, this._uninstallCacheLocalPaths, backupTopPath);
                if (result != true)
                {
                    MessageBox.Show("ユーザーキャッシュの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("ユーザーキャッシュを削除しました。");
                }
            }

            /* ユーザーのレジストリデータを削除する */
            if (checkBoxUninstallDeleteUserRegistory.Checked == true)
            {
                bool result = await deleteAndExportUserRegistory(ironcadPaths, this._uninstallUserRegPaths, backupTopPath);
                if (result != true)
                {
                    MessageBox.Show("ユーザーレジストリの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("ユーザーレジストリを削除しました。");
                }
            }

            /* 端末のレジストリデータを削除する */
            if (checkBoxUninstallDeleteMachineRegistory.Checked == true)
            {
                bool result = await deleteAndExportMachineRegistory(ironcadPaths, this._uninstallMachineRegPaths, backupTopPath);
                if (result != true)
                {
                    MessageBox.Show("ベース レジストリの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("ベース レジストリを削除しました。");
                }
            }

            /* 端末のIRONCADフォルダを削除する */
            if (checkBoxUninstallDeleteIRONCADFolder.Checked == true)
            {
                bool result = await deleteIRONCADFolder(ironcadPaths, this._uninstallInstallPaths);
                if (result != true)
                {
                    MessageBox.Show("IRONCADインストールフォルダの削除中にエラーが発生しました。");
                }
                else
                {
                    MessageBox.Show("IRONCADインストールフォルダを削除しました。");
                }
            }

            ((UserControlTagData)this.Tag).canNotClose = false;
            buttonUninstallDelete.Enabled = true;
            initializeControlAndInnerData(comboBoxIroncadVersion.Text);
        }

        /// <summary>
        /// 削除するユーザーキャッシュの情報を取得する
        /// </summary>
        /// <param name="checkbox">チェックボックス</param>
        /// <param name="isUninstall">true: アンインストール後の削除  false:ユーザーデータの削除</param>
        /// <param name="sb">out:削除情報</param>
        private void getUserCacheDeleteInformation(CheckBox checkbox, bool isUninstall, ref StringBuilder sb)
        {
            if (checkbox.Checked == true)
            {
                sb.AppendLine("***** 下記パスのユーザキャッシュデータを削除します *****");
                if(isUninstall == true)
                {
                    foreach (string path in this._uninstallCacheRoamingPaths)
                    {
                        sb.AppendLine(path);
                    }
                    foreach (string path in this._uninstallCacheLocalPaths)
                    {
                        sb.AppendLine(path);
                    }
                }
                else
                {
                    foreach (string path in this._workCacheRoamingPaths)
                    {
                        sb.AppendLine(path);
                    }
                    foreach (string path in this._workCacheLocalPaths)
                    {
                        sb.AppendLine(path);
                    }
                }
            }
        }


        /// <summary>
        /// 削除するユーザーレジストリの情報を取得する
        /// </summary>
        /// <param name="checkbox">チェックボックス</param>
        /// <param name="isUninstall">true: アンインストール後の削除  false:ユーザーデータの削除</param>
        /// <param name="sb">out:削除情報</param>
        private void getUserRegDeleteInformation(CheckBox checkbox, bool isUninstall, ref StringBuilder sb)
        {
            if (checkbox.Checked == true)
            {
                sb.AppendLine("***** 下記パスのユーザレジストリデータを削除します *****");
                if (isUninstall == true)
                {
                    foreach (string regPath in this._uninstallUserRegPaths)
                    {
                        sb.AppendLine(Path.Combine(@"\HKEY_CURRENT_USER", regPath));
                    }
                }
                else
                {
                    foreach (string regPath in this._workUserRegPaths)
                    {
                        sb.AppendLine(Path.Combine(@"\HKEY_CURRENT_USER", regPath));
                    }
                }
            }
        }


        /// <summary>
        /// 削除する端末レジストリの情報を取得する
        /// </summary>
        /// <param name="sb">out:削除情報</param>
        private void getMachineRegDeleteInformation(ref StringBuilder sb)
        {
            if (checkBoxUninstallDeleteMachineRegistory.Checked == true)
            {
                sb.AppendLine("***** 下記パスの端末レジストリデータを削除します *****");
                foreach (string regPath in this._uninstallMachineRegPaths)
                {
                    sb.AppendLine(Path.Combine(@"\HKEY_LOCAL_MACHINE", regPath));
                }
            }
        }


        /// <summary>
        /// 削除するIRONCADフォルダの情報を取得する
        /// </summary>
        /// <param name="sb">out:削除情報</param>
        private void getInstallDirDeleteInformation(ref StringBuilder sb)
        {
            if(checkBoxUninstallDeleteIRONCADFolder.Checked == true)
            {
                sb.AppendLine("***** 下記パスのIRONCADインストールデータを削除します *****");
                foreach(string path in this._uninstallInstallPaths)
                {
                    sb.AppendLine(path);
                }
            }
        }


        /// <summary>
        /// 削除するパス情報とレジストリ情報を更新する
        /// </summary>
        private void refreshDeletePathAndRegList()
        {
            this._workCacheLocalPaths = new List<string>();
            this._workCacheRoamingPaths = new List<string>();
            this._uninstallCacheLocalPaths = new List<string>();
            this._uninstallCacheRoamingPaths = new List<string>();
            this._workUserRegPaths = new List<string>();
            this._uninstallUserRegPaths = new List<string>();
            this._uninstallInstallPaths = new List<string>();
            this._uninstallMachineRegPaths = new List<string>();
            IRONCAD_REG_AND_DIR ironcadPaths = getSelectIRONCADVersion();
            if (ironcadPaths == null)
            {
                return;
            }
            /**************************************************/
            /* ユーザのデータ (WorkDeleteとUninstallDelete))  */
            /**************************************************/
            /* Local */
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCOMPOSE) != true) &&
               (Directory.Exists(ironcadPaths.cacheDirLocalCOMPOSE) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCOMPOSE);
                if(this.checkBoxWorkDeleteCacheCOMPOSE.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCOMPOSE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalDRAFT) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirLocalDRAFT) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalDRAFT);
                if(this.checkBoxWorkDeleteCacheDRAFT.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalDRAFT);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalINOVATE) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirLocalINOVATE) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalINOVATE);
                if (this.checkBoxWorkDeleteCacheINOVATE.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalINOVATE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalIRONCAD) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirLocalIRONCAD) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalIRONCAD);
                if (this.checkBoxWorkDeleteCacheIRONCAD.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalIRONCAD);
                }
            }


            /* LocalCaxa */
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaIC) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaIC) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIC);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIC);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaIntDRAFT) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaIntDRAFT) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntDRAFT);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntDRAFT);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaIntINOVATE) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaIntINOVATE) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntINOVATE);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntINOVATE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaIntIC) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaIntIC) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntIC);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaIntIC);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaPrintTool) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaPrintTool) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaPrintTool);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaPrintTool);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirLocalCaxaExbViewr) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirLocalCaxaExbViewr) == true))
            {
                this._uninstallCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaExbViewr);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheLocalPaths.Add(ironcadPaths.cacheDirLocalCaxaExbViewr);
                }
            }

            /* Roaming */
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCOMPOSE) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirRoamingCOMPOSE) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCOMPOSE);
                if (this.checkBoxWorkDeleteCacheCOMPOSE.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCOMPOSE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingDRAFT) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirRoamingDRAFT) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingDRAFT);
                if (this.checkBoxWorkDeleteCacheDRAFT.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingDRAFT);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingINOVATE) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirRoamingINOVATE) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingINOVATE);
                if (this.checkBoxWorkDeleteCacheINOVATE.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingINOVATE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingIRONCAD) != true) &&
                (Directory.Exists(ironcadPaths.cacheDirRoamingIRONCAD) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingIRONCAD);
                if (this.checkBoxWorkDeleteCacheIRONCAD.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingIRONCAD);
                }
            }


            /* RoamingCaxa */
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaIC) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaIC) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIC);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIC);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaIntDRAFT) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaIntDRAFT) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntDRAFT);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntDRAFT);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaIntINOVATE) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaIntINOVATE) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntINOVATE);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntINOVATE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaIntIC) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaIntIC) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntIC);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaIntIC);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaPrintTool) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaPrintTool) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaPrintTool);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaPrintTool);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.cacheDirRoamingCaxaExbViewr) != true) &&
                 (Directory.Exists(ironcadPaths.cacheDirRoamingCaxaExbViewr) == true))
            {
                this._uninstallCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaExbViewr);
                if (this.checkBoxWorkDeleteCacheCaxaDraft.Checked == true)
                {
                    this._workCacheRoamingPaths.Add(ironcadPaths.cacheDirRoamingCaxaExbViewr);
                }
            }


            /******************************************************/
            /* ユーザのレジストリ (WorkDeleteとUninstallDelete))  */
            /******************************************************/
            if ((string.IsNullOrEmpty(ironcadPaths.regPathUserCOMPOSE) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathUserCOMPOSE, false)) != true))
            {
                this._uninstallUserRegPaths.Add(ironcadPaths.regPathUserCOMPOSE);
                if (this.checkBoxWorkDeleteCacheCOMPOSE.Checked == true)
                {
                    this._workUserRegPaths.Add(ironcadPaths.regPathUserCOMPOSE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathUserDRAFT) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathUserDRAFT, false)) != true))
            {
                this._uninstallUserRegPaths.Add(ironcadPaths.regPathUserDRAFT);
                if (this.checkBoxWorkDeleteCacheDRAFT.Checked == true)
                {
                    this._workUserRegPaths.Add(ironcadPaths.regPathUserDRAFT);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathUserINOVATE) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathUserINOVATE, false)) != true))
            {
                this._uninstallUserRegPaths.Add(ironcadPaths.regPathUserINOVATE);
                if (this.checkBoxWorkDeleteCacheINOVATE.Checked == true)
                {
                    this._workUserRegPaths.Add(ironcadPaths.regPathUserINOVATE);
                }
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathUserIRONCAD) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathUserIRONCAD, false)) != true))
            {
                this._uninstallUserRegPaths.Add(ironcadPaths.regPathUserIRONCAD);
                if (this.checkBoxWorkDeleteCacheIRONCAD.Checked == true)
                {
                    this._workUserRegPaths.Add(ironcadPaths.regPathUserIRONCAD);
                }
            }

            /**************************************/
            /* 端末全体のデータ (UninstallDelete) */
            /**************************************/
            /* IRONCADフォルダ */
            if ((string.IsNullOrEmpty(ironcadPaths.installDir) != true) &&
                (Directory.Exists(ironcadPaths.installDir) == true))
            {
                this._uninstallInstallPaths.Add(ironcadPaths.installDir);
            }

            /* Baseレジストリ */
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineCOMPOSE) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineCOMPOSE, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineCOMPOSE);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineCOMPOSE_NewUser) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineCOMPOSE_NewUser, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineCOMPOSE_NewUser);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineDRAFT) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineDRAFT, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineDRAFT);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineDRAFT_NewUser) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineDRAFT_NewUser, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineDRAFT_NewUser);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineINOVATE) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineINOVATE, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineINOVATE);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineINOVATE_NewUser) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineINOVATE_NewUser, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineINOVATE_NewUser);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineIRONCAD) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineIRONCAD, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineIRONCAD);
            }
            if ((string.IsNullOrEmpty(ironcadPaths.regPathMachineIRONCAD_NewUser) != true) &&
                (string.IsNullOrEmpty(icapiCommon.existsRegistoryPath(ironcadPaths.regPathMachineIRONCAD_NewUser, true)) != true))
            {
                this._uninstallMachineRegPaths.Add(ironcadPaths.regPathMachineIRONCAD_NewUser);
            }

            showWorkDeleteInformation();
            showUninstallDeleteInformation();
        }


        /// <summary>
        /// 削除対象のパス・レジストリを表示する(アンインストール後の削除)
        /// </summary>
        private void showUninstallDeleteInformation()
        {
            StringBuilder sb = new StringBuilder();
            getUserCacheDeleteInformation(checkBoxUninstallDeleteCache, true, ref sb);
            getUserRegDeleteInformation(checkBoxUninstallDeleteUserRegistory, true, ref sb);
            getMachineRegDeleteInformation(ref sb);
            getInstallDirDeleteInformation(ref sb);
            textBoxUninstallDelete.Text = sb.ToString();
        }


        /// <summary>
        /// 削除対象のパス・レジストリを表示する(ユーザーキャッシュの削除)
        /// </summary>
        private void showWorkDeleteInformation()
        {
            StringBuilder sb = new StringBuilder();
            getUserCacheDeleteInformation(checkBoxWorkDeleteCache, false, ref sb);
            getUserRegDeleteInformation(checkBoxWorkDeleteUserRegistory, false, ref sb);
            textBoxWorkDelete.Text = sb.ToString();
        }


        #region イベント
        /// <summary>
        /// バージョン選択のcomboBox変更イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxIroncadVersion_TextChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
            return;
        }


        /// <summary>
        /// CheckBox(ユーザーキャッシュの削除)のチェック状態変更イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCache_CheckedChanged(object sender, EventArgs e)
        {
            showWorkDeleteInformation();
        }


        /// <summary>
        /// CheckBox(ユーザーレジストリの削除)のチェック状態変更イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteUserRegistory_CheckedChanged(object sender, EventArgs e)
        {
            showWorkDeleteInformation();
        }

        /// <summary>
        /// CheckBox(IRONCAD)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCacheIRONCAD_CheckedChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
        }

        /// <summary>
        /// CheckBox(2D ドラフト(Caxa Draft))のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCacheCaxaDraft_CheckedChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
        }


        /// <summary>
        /// CheckBox(INOVATE)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCacheINOVATE_CheckedChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
        }


        /// <summary>
        /// CheckBox(COMPOSE)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCacheCOMPOSE_CheckedChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
        }


        /// <summary>
        /// CheckBox(DRAFT)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxWorkDeleteCacheDRAFT_CheckedChanged(object sender, EventArgs e)
        {
            refreshDeletePathAndRegList();
        }


        /// <summary>
        /// CheckBox(ユーザーキャッシュの削除(上と同じ))のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxUninstallDeleteCache_CheckedChanged(object sender, EventArgs e)
        {
            showUninstallDeleteInformation();
        }


        /// <summary>
        /// CheckBox(ユーザーレジストリの削除(上と同じ))のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxUninstallDeleteUserRegistory_CheckedChanged(object sender, EventArgs e)
        {
            showUninstallDeleteInformation();
        }


        /// <summary>
        /// CheckBox(ベースレジストリの削除※)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxUninstallDeleteMachineRegistory_CheckedChanged(object sender, EventArgs e)
        {
            showUninstallDeleteInformation();
        }


        /// <summary>
        /// CheckBox(IRONCADフォルダの削除※)のチェック状態変更イベント 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxUninstallDeleteIRONCADFolder_CheckedChanged(object sender, EventArgs e)
        {
            showUninstallDeleteInformation();
        }


        #endregion イベント
    }
}
