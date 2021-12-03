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
using Microsoft.Win32;

namespace ICApiAddin.icMaintenanceTools
{
    public partial class UserControlAllIroncadRepair: UserControl
    {
        public const string title = "全バージョンのIRONCADの修復";

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public UserControlAllIroncadRepair()
        {
            InitializeComponent();
            this.Tag = new UserControlTagData();
            this.Dock = DockStyle.Fill;
        }


        /// <summary>
        /// ねじ山画像のファイルがあるかチェックする
        /// </summary>
        /// <param name="threadImageFilePath">見つかったねじ山画像のファイルパス</param>
        /// <returns></returns>
        private bool checkThreadFile(ref string threadImageFilePath)
        {
            threadImageFilePath = string.Empty;
            string threadFileValue = string.Empty;
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\IronCAD\InnovationSuite\Directories\");
            if (rkey == null)
            {
                return false;
            }
            try
            {
                threadFileValue = rkey.GetValue("ThreadFile").ToString();
                if (string.IsNullOrEmpty(threadFileValue) == true)
                {
                    rkey.Close();
                    return false;
                }
            }
            catch(Exception ex)
            {
                rkey.Close();
                return false;
            }
            finally
            {
                rkey.Close();
            }

            if (Directory.Exists(threadFileValue) != true)
            {
                return false;
            }

            string tmpThreadImageFilePath = Path.Combine(threadFileValue, "3iaoftt.jpg");
            if(File.Exists(tmpThreadImageFilePath) != true)
            {
                return false;
            }
            threadImageFilePath = tmpThreadImageFilePath;

            return true;
        }


        /// <summary>
        /// IRONCADのプログラムフォルダからねじ山画像を検索する
        /// </summary>
        /// <param name="threadFilePath">見つかったねじ山画像のファイルパス</param>
        /// <returns></returns>
        private bool searchThreadFileFromIRONCAD(ref string threadFilePath)
        {
            threadFilePath = string.Empty;
            bool fileExists = false;
            List<KeyValuePair<string, string>> allIroncad = new List<KeyValuePair<string, string>>();
            icapiCommon.GetAllIronCADInstallPath(ref allIroncad, true, true);
            if (allIroncad.Count() <= 0)
            {
                return false;
            }

            foreach (KeyValuePair<string, string> item in allIroncad)
            {
                string ironcadPath = item.Value;
                if (Directory.Exists(ironcadPath) != true)
                {
                    continue;
                }
                string checkThreadFilePath = Path.Combine(ironcadPath, "bin\\3iaoftt.jpg");
                if (File.Exists(checkThreadFilePath) == true)
                {
                    fileExists = true;
                    threadFilePath = checkThreadFilePath;
                    break;
                }
            }
            return fileExists;
        }


        /// <summary>
        /// IRONCADのプログラムフォルダにねじ山画像を生成する
        /// </summary>
        /// <param name="threadFilePath">生成したねじ山画像のファイルパス</param>
        /// <returns></returns>
        private bool createThreadFileToIRONCAD(ref string threadFilePath)
        {
            bool writeResult = false;
            threadFilePath = string.Empty;
            List<KeyValuePair<string, string>> allIroncad = new List<KeyValuePair<string, string>>();
            icapiCommon.GetAllIronCADInstallPath(ref allIroncad, true, true);
            if (allIroncad.Count() <= 0)
            {
                return false;
            }

            foreach (KeyValuePair<string, string> item in allIroncad)
            {
                string ironcadBinPath = Path.Combine(item.Value, "bin");
                if (Directory.Exists(ironcadBinPath) != true)
                {
                    continue;
                }
                string writeThreadFilePath = Path.Combine(ironcadBinPath, "3iaoftt.jpg");
                try
                {
                    byte[] threadData = Properties.Resources._3iaoftt;
                    System.IO.File.WriteAllBytes(writeThreadFilePath, threadData);
                    threadFilePath = writeThreadFilePath;
                    writeResult = true;
                    break;
                }
                catch(Exception ex)
                {
                    writeResult = false;
                }
            }
            return writeResult;
        }


        /// <summary>
        /// ねじ山のレジストリを作成する
        /// </summary>
        /// <param name="threadFilePath">ねじ山画像のファイルパス</param>
        /// <returns></returns>
        private bool createRegistoryThreadFile(string threadFilePath)
        {
            bool result = false;
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\IronCAD\InnovationSuite\Directories\", true);
            if (rkey == null)
            {
                return false;
            }
            try
            {
                string threadFileValue = rkey.GetValue("ThreadFile").ToString();
                if (string.IsNullOrEmpty(threadFileValue) != true)
                {
                    rkey.DeleteValue("ThreadFile");
                }
            }
            catch(Exception ex)
            {

            }
            try
            {
                rkey.SetValue("ThreadFile", Path.GetDirectoryName(threadFilePath));
                result = true;
            }
            catch(Exception ex)
            {
                result = false;
            }
            rkey.Close();

            return result;
        }


        /// <summary>
        /// IMEの設定変更（旧IMEを使用するか新IMEを使用するか）
        /// </summary>
        /// <param name="useOldIME">true:旧IMEを使用 false:新IMEを使用</param>
        /// <returns></returns>
        public bool setOldIMEversion(bool useOldIME)
        {
            bool result = false;
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Input\TSF\Tsf3Override\{03b5835f-f03c-411b-9ce2-aa23e1171e36}", true);
            if (rkey == null)
            {
                return false;
            }
            try
            {
                string threadFileValue = rkey.GetValue("NoTsf3Override2").ToString();
                if (string.IsNullOrEmpty(threadFileValue) != true)
                {
                    rkey.DeleteValue("NoTsf3Override2");
                }
            }
            catch (Exception ex)
            {

            }
            try
            {
                if (useOldIME == true)
                {
                    rkey.SetValue("NoTsf3Override2", 1);
                }
                else
                {
                    rkey.SetValue("NoTsf3Override2", 0);
                }
                result = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            rkey.Close();

            return result;
        }


        /// <summary>
        /// CADENASのワークデータを削除する（バックアップもする）
        /// </summary>
        /// <param name="cadenasWorkPath"></param>
        /// <param name="backupTopPath"></param>
        /// <returns></returns>
        private async Task<bool> backupAndDeleteCadenasWork(string cadenasWorkPath, string backupTopPath)
        {
            string appDataRoamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            if ((string.IsNullOrEmpty(cadenasWorkPath) == false &&
                (Directory.Exists(cadenasWorkPath) == true)))
            {
                string r_baseFolderName = Path.GetFileName(appDataRoamingPath); /* Roaming */
                string r_prgName = cadenasWorkPath.Replace(appDataRoamingPath, string.Empty).Trim('\\'); /* cadenas */
                string r_backupPath = Path.Combine(backupTopPath, r_baseFolderName, r_prgName);

                bool r_result = await icapiCommon.BackupDirectory(cadenasWorkPath, r_backupPath);
                if (r_result != true)
                {
                    return false;
                }
                r_result = await icapiCommon.deleteDirectory(cadenasWorkPath);
                if (r_result != true)
                {
                    return false;
                }
            }
            return true;
        }


        /// <summary>
        /// フォント設定のショートカットを使用したフォントのインストールを許可する(上級者用)を操作する
        /// </summary>
        /// <param name="setValue">設定値</param>
        /// <param name="oldValue">設定前の値</param>
        /// <returns></returns>
        private bool setFontInstallAsLink(int setValue, ref int oldValue)
        {
            bool ret = false;
            oldValue = -1;
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Font Management", true);
            if (rkey == null)
            {
                return false;
            }
            try
            {
                oldValue = (int)rkey.GetValue("InstallAsLink");

                /* shortcutでのインストールを許可する */
                rkey.SetValue("InstallAsLink", setValue, RegistryValueKind.DWord);
                ret = true;
            }
            catch (Exception ex)
            {
                rkey.Close();
                ret = false;
            }
            finally
            {
                rkey.Close();
            }
            return ret;
        }


        /// <summary>
        /// GDTフォント設定を修復する
        /// </summary>
        private bool RepairGDTFontSetting()
        {
            List<KeyValuePair<string, string>> ironcadList = new List<KeyValuePair<string, string>>();
            icapiCommon.GetAllIronCADInstallPath(ref ironcadList, true, true);
            string setFontGDT_Path = string.Empty;
            string setFontGDTSHP_Path = string.Empty;
            foreach (KeyValuePair<string, string> ironcad in ironcadList)
            {
                string installPath = ironcad.Value;
                string fontPath = Path.Combine(installPath, @"bin\CAXADraft\Font");
                string fontGDT_Path = Path.Combine(installPath, @"bin\CAXADraft\Font\CXGDT.ttf");
                string fontGDTSHP_Path = Path.Combine(installPath, @"bin\CAXADraft\Font\CXGDTSHP.ttf");
                if ((File.Exists(fontGDT_Path) == true) &&
                    (File.Exists(fontGDTSHP_Path) == true))
                {
                    setFontGDT_Path = fontGDT_Path;
                    setFontGDTSHP_Path = fontGDTSHP_Path;
                    break;
                }
            }
            if ((string.IsNullOrEmpty(setFontGDT_Path) == true) ||
                (string.IsNullOrEmpty(setFontGDTSHP_Path) == true))
            {
                MessageBox.Show("GDTフォントが見つかりませんでした。");
                return false;
            }
            bool result = setRegistoryGDTFont(setFontGDT_Path, setFontGDTSHP_Path);
            return result;
        }


        /// <summary>
        /// GDTフォントを追加する
        /// </summary>
        /// <param name="fontGDT_Path"></param>
        /// <param name="fontGDTSHP_Path"></param>
        /// <returns></returns>
        public bool setRegistoryGDTFont(string fontGDT_Path, string fontGDTSHP_Path)
        {
            bool result = false;
            Microsoft.Win32.RegistryKey rkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", true);
            if (rkey == null)
            {
                return false;
            }
            try
            {
                rkey.SetValue("CXGDT (TrueType)", fontGDT_Path, RegistryValueKind.String);
                rkey.SetValue("CXGDTSHP (TrueType)", fontGDTSHP_Path, RegistryValueKind.String);
                result = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                rkey.Close();
            }

            return result;
        }

        #region イベント
        /// <summary>
        /// ねじ山画像表示の機能修復ボタン クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRepairThread_Click(object sender, EventArgs e)
        {
            string threadFilePath = string.Empty;
            bool checkResult = checkThreadFile(ref threadFilePath);
            if(checkResult == true)
            {
                DialogResult dret = MessageBox.Show("問題が見つかりませんでした。修復を強制的に実行しますか？", "確認", MessageBoxButtons.YesNo);
                if(dret != DialogResult.Yes)
                {
                    return;
                }
            }

            bool exists = searchThreadFileFromIRONCAD(ref threadFilePath);
            if (exists != true)
            {
                bool ret = createThreadFileToIRONCAD(ref threadFilePath);
                if(ret != true)
                {
                    return;
                }
            }
            createRegistoryThreadFile(threadFilePath);
            MessageBox.Show("修復が完了しました。\nIRONCADを起動している場合は再起動してください。");
        }


        /// <summary>
        /// IMEによるIRONCADフリーズを回避(修復)するボタン クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonUseOldIME_Click(object sender, EventArgs e)
        {
            bool ret = setOldIMEversion(true);
            if(ret != true)
            {
                MessageBox.Show("実行中にエラーが発生しました。");
                return;
            }
            MessageBox.Show("修復が完了しました。\nIRONCADを起動している場合は再起動してください。");
        }


        /// <summary>
        ///  IMEによるIRONCADフリーズを回避(修復)を戻すラベル クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabelUseNewIME_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            bool ret = setOldIMEversion(false);
            if (ret != true)
            {
                MessageBox.Show("実行中にエラーが発生しました。");
                return;
            }
            MessageBox.Show("設定を変更しました。\nIRONCADを起動している場合は再起動してください。");
        }


        /// <summary>
        /// GDTフォント設定を修復するボタン クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRepairGDTFont_Click(object sender, EventArgs e)
        {
            bool ret = RepairGDTFontSetting();
            if (ret != true)
            {
                MessageBox.Show("実行中にエラーが発生しました。");
                return;
            }
            MessageBox.Show("修復が完了しました。\nIRONCADを起動している場合は再起動してください。");
        }


        /// <summary>
        /// CADENASのワークデータを削除する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void buttonRepairCadenas_Click(object sender, EventArgs e)
        {
            string appDataRoamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string cadenasWorkPath = Path.Combine(appDataRoamingPath, "cadenas");
            string backupTopPath = "backup" + DateTime.Now.ToString("yyyyMMddHHmmss");

            bool result = await backupAndDeleteCadenasWork(cadenasWorkPath, backupTopPath);
            if (result == true)
            {
                MessageBox.Show("CADENASのワークデータを削除しました。\nIRONCADを起動している場合は再起動してください。");
            }
        }

        #endregion イベント
    }
}
