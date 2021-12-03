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
using System.Net;
using System.Net.Http;

namespace ICApiAddin.icMaintenanceTools
{
    public partial class UserControlUpdateList : UserControl
    {
        public const string title = "IRONCADの更新プログラム(公式)";
        /// <summary>
        /// IRONCAD公式のアップデート情報
        /// </summary>
        private class IRONCAD_UPDATE
        {
            public string title;
            public string date;
            public string link;
            public string about;
            public IRONCAD_UPDATE(string title, string date, string link)
            {
                this.title = title;
                this.date = date;
                this.link = link;
                this.about = string.Empty;
            }
            public IRONCAD_UPDATE(string title, string date, string link, string about)
            {
                this.title = title;
                this.date = date;
                this.link = link;
                this.about = about;
            }
            public IRONCAD_UPDATE()
            {
                this.title = string.Empty;
                this.date = string.Empty;
                this.link = string.Empty;
                this.about = string.Empty;
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public UserControlUpdateList()
        {
            InitializeComponent();
            this.Tag = new UserControlTagData();
            this.Dock = DockStyle.Fill;
            object[] obj = new object[] { "更新をチェック中です。しばらくお待ちください。" };
            this.dataGridViewUpdateList.Rows.Add(obj);
        }

        /// <summary>
        /// ページ表示状態変更イベント(ページ表示イベントとして利用)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void UserControlUpdateList_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible && !Disposing)
            {
                ((UserControlTagData)this.Tag).canNotClose = true;
                this.Cursor = Cursors.WaitCursor;
                this.dataGridViewUpdateList.Cursor = this.Cursor;
                bool result = await ShowIroncadUpdateList();
                this.Cursor = Cursors.Default;
                this.dataGridViewUpdateList.Cursor = this.Cursor;
                ((UserControlTagData)this.Tag).canNotClose = false;
            }
        }

        private async Task<bool> ShowIroncadUpdateList()
        {
            List<IRONCAD_UPDATE> updateList = new List<IRONCAD_UPDATE>();
            updateList = await getUpdateList();
            this.dataGridViewUpdateList.Rows.Clear();
            foreach (IRONCAD_UPDATE update in updateList)
            {
                DataGridViewRow dr = new DataGridViewRow();
                dr.CreateCells(this.dataGridViewUpdateList);
                object[] obj = new object[] { update.date, update.title, update.about };
                dataGridViewUpdateList.Rows.Add(obj);
                int addedRowIndex = dataGridViewUpdateList.Rows.Count - 1;
                dataGridViewUpdateList["UpdateTitle", addedRowIndex].Tag = update.link;
            }
            return true;
        }



        private async Task<List<IRONCAD_UPDATE>> getUpdateList()
        {
            List<IRONCAD_UPDATE> updateList = new List<IRONCAD_UPDATE>();
            Uri webUri = new Uri("https://www.ironcad.jp/product-updates/");
            string htmlText = await GetPage(webUri);

            if (htmlText != null)
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(htmlText);
                HtmlAgilityPack.HtmlNodeCollection nodes = htmlDoc.DocumentNode.SelectNodes(@"//div[@class=""card-content""]");
                foreach (HtmlAgilityPack.HtmlNode node in nodes)
                {
                    try
                    {
                        var title = node.ChildNodes["h3"].InnerHtml.Trim();
                        var date = node.ChildNodes["p"].InnerHtml.Trim();
                        var link = node.ChildNodes.Where(a => a.InnerHtml.Contains("btn btn-default btn-primary btn-cst") == true).FirstOrDefault().ChildNodes["a"].Attributes["href"].Value;
                        //.SelectSingleNode(@"//h3[@class=""headline""]").InnerHtml.Trim();
                        //                        string link = node.SelectSingleNode(@"//div[@class=""btn-container""]/a[@class=""btn btn-default btn-primary btn-cstm""]").Attributes["href"].Value;
                        //string link = node.ChildNodes..Select(a => a.Attributes. ["href"].Value);
                        updateList.Add(new IRONCAD_UPDATE(title, date, link));
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            for(int i = 0; i < updateList.Count; i++)
            {
                string link = updateList[i].link;
                Uri aboutUri = new Uri(link);
                string aboutHtmlText = await GetPage(aboutUri);
                {
                    if (aboutHtmlText != null)
                    {
                        var aboutHtmlDoc = new HtmlAgilityPack.HtmlDocument();
                        aboutHtmlDoc.LoadHtml(aboutHtmlText);
                        try
                        {
                            var about = aboutHtmlDoc.DocumentNode.SelectNodes(@"//meta[@property=""og:description""]").FirstOrDefault().Attributes["content"].Value;
                            updateList[i].about = about.Replace("... 全文を表示", "....." + Environment.NewLine + "(詳しくは左のタイトルをクリックして公式サイトを確認ください)");
                        }
                        catch(Exception ex)
                        {

                        }
                    }
                }
            }
            return updateList;
        }

        private async Task<string> GetPage(Uri uri)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko");
                client.DefaultRequestHeaders.Add("Accept-Language", "ja-JP");
                client.Timeout = TimeSpan.FromSeconds(10.0);
                try
                {
                    return await client.GetStringAsync(uri);
                }
                catch (Exception e)
                {

                }
                return null;
            }
        }


        private void linkLabelIroncadUpdate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url = "https://www.ironcad.jp/product-updates/";
            System.Diagnostics.Process.Start(url);
        }

        private void dataGridViewUpdateList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int columnIndex = e.ColumnIndex;
            int rowIndex = e.RowIndex;
            int linkColumnIndex = dataGridViewUpdateList.Columns["UpdateTitle"].Index;
            if(columnIndex == linkColumnIndex)
            {
                string url = dataGridViewUpdateList[columnIndex, rowIndex].Tag.ToString();
                if(string.IsNullOrEmpty(url) != true)
                {
                    System.Diagnostics.Process.Start(url);
                }
            }

        }
    }
}
