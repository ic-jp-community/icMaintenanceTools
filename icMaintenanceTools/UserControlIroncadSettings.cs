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
using Microsoft.Win32;

namespace ICApiAddin.icMaintenanceTools
{
    public partial class UserControlIroncadSettings : UserControl
    {
        public const string title = "IRONCADの設定";
        public List<IRONCAD_REG_AND_DIR> _ironcadList = new List<IRONCAD_REG_AND_DIR>();
        public UserControlIroncadSettings()
        {
            InitializeComponent();
            this.Tag = new UserControlTagData();
            this.Dock = DockStyle.Fill;
        }

        private void UserControlIroncadRepair_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible && !Disposing)
            {
                this.Cursor = Cursors.WaitCursor;
                getIRONCADInformation(string.Empty);
                this.Cursor = Cursors.Default;
            }
        }

        private string langCodeToLangStr(int code)
        {
            string lang = string.Empty;
            switch (code)
            {
                case 0:
                    lang = "英語(0)";
                    break;
                case 1041:
                    lang = "日本語(1041)";
                    break;
                default:
                    lang = "不明(" + code.ToString() + ")";
                    break;
            }
            return lang;
        }

        private int langStrToLangCode(string langStr)
        {
            int langCode = -1;
            switch (langStr)
            {
                case "英語(0)":
                    langCode = 0;
                    break;
                case "日本語(1041)":
                    langCode = 1041;
                    break;
                default:
                    langCode = -1;
                    break;
            }
            return langCode;
        }
        private string langNameToLangStr(string name)
        {
            string lang = string.Empty;
            string name_toLower = name.ToLower();
            switch (name_toLower)
            {
                case "en-us":
                    lang = "英語(0)";
                    break;
                case "ja-jp":
                    lang = "日本語(1041)";
                    break;
                default:
                    lang = "不明(-)";
                    break;
            }
            return lang;
        }
        private void GetCurrentLanguage(string version)
        {
            string numVersionStr = icapiCommon.convertYearVersionToNumberVersion(Int32.Parse(version));
            double subVersion = double.Parse(numVersionStr);
            string regPath = @"Software\IronCAD" + @"\IRONCAD " + subVersion.ToString("F1") + @"\General";
            RegistryKey reg = Registry.CurrentUser.OpenSubKey(regPath);
            object obj = reg.GetValue("Current Language");
            reg.Close();

            textBoxCurrLang.Text = langCodeToLangStr(Int32.Parse(obj.ToString()));

        }
        private void SetCurrentLanguage(string version, int code)
        {
            string numVersionStr = icapiCommon.convertYearVersionToNumberVersion(Int32.Parse(version));
            double subVersion = double.Parse(numVersionStr);
            string regPath = @"Software\IronCAD" + @"\IRONCAD " + subVersion.ToString("F1") + @"\General";
            RegistryKey reg = Registry.CurrentUser.CreateSubKey(regPath);
            reg.SetValue("Current Language", code);
            reg.Close();
        }
        private void getInstallLanguage(string version, ref List<string> langList)
        {
            string numVersionStr = icapiCommon.convertYearVersionToNumberVersion(Int32.Parse(version));
            double subVersion = double.Parse(numVersionStr);
            string regPath = @"Software\IronCAD" + @"\IRONCAD " + subVersion.ToString("F1") + @"\LangDependent";
            RegistryKey reg = Registry.CurrentUser.OpenSubKey(regPath);
            string[] arySubKeyNames = reg.GetSubKeyNames();
            reg.Close();
            foreach (string str in arySubKeyNames)
            {
                langList.Add(langNameToLangStr(str));
            }
        }

        private void getIRONCADInformation(string selectVersion)
        {
            this._ironcadList.Clear();
            this._ironcadList = icapiCommon.getAllIroncadInformation();

            comboBoxIroncadVersion.DataSource = this._ironcadList;
            comboBoxIroncadVersion.DisplayMember = "version";
            comboBoxIroncadVersion.ValueMember = "installDir";

            if(string.IsNullOrEmpty(selectVersion) != true) 
            {
                IEnumerable<IRONCAD_REG_AND_DIR> found = this._ironcadList.Where(a => string.Equals(a.version, selectVersion));
                if(found != null)
                {
                    comboBoxIroncadVersion.Text = selectVersion;
                }
            }
        }

        private string existsRegistoryPath(string regPath, bool isMachine)
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

        private string existDirectoryPath(string dirPath)
        {
            if(Directory.Exists(dirPath) != true)
            {
                return string.Empty;
            }
            return dirPath;
        }

        private IRONCAD_REG_AND_DIR getSelectIRONCADVersion()
        {
            string version = comboBoxIroncadVersion.Text;
            IRONCAD_REG_AND_DIR data = this._ironcadList.Where(a => a.version == version).FirstOrDefault();
            return data;
        }

        private void RefreshLanguage()
        {
            IRONCAD_REG_AND_DIR data = getSelectIRONCADVersion();
            if(data == null)
            {
                return;
            }
            string ironCadVersion = data.version;
            List<string> LangList = new List<string>();
            GetCurrentLanguage(ironCadVersion);
            getInstallLanguage(ironCadVersion, ref LangList);
            comboBoxSetLanguage.Items.Clear();
            foreach(string str in LangList)
            {
                comboBoxSetLanguage.Items.Add(str);
            }
            comboBoxSetLanguage.SelectedItem = comboBoxSetLanguage.Items[0];
        }
        private void comboBoxIroncadVersion_TextChanged(object sender, EventArgs e)
        {
            RefreshLanguage();
            return;
        }

        private void buttonSetLanguage_Click(object sender, EventArgs e)
        {
            IRONCAD_REG_AND_DIR data = getSelectIRONCADVersion();
            string ironCadVersion = data.version;

            if (string.IsNullOrEmpty(comboBoxSetLanguage.SelectedItem.ToString()) == true)
            {
                MessageBox.Show("変更後の言語を指定してください。");
                return;
            }
            int code = langStrToLangCode(comboBoxSetLanguage.SelectedItem.ToString());
            if (code == -1)
            {
                MessageBox.Show("変更後の言語が不正な値です。\n変更後の言語を確認してください。");
                return;
            }

            DialogResult ret = MessageBox.Show(comboBoxSetLanguage.SelectedItem.ToString() + " に表示言語を変更します。\nよろしいですか？", "確認", MessageBoxButtons.OKCancel);
            if (ret != DialogResult.OK)
            {
                return;
            }

            SetCurrentLanguage(ironCadVersion, code);
            MessageBox.Show("表示言語の設定を変更しました。IronCADを再起動してください。");
            GetCurrentLanguage(ironCadVersion);
        }
    }
}
