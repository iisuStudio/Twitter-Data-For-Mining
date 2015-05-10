using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//新增的命名空間
using System.Net;
using System.IO;
using TweetSharp;
using Newtonsoft.Json;
using System.Data.Odbc;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {

        TwitterService servis;
        public Form1()
        {
            InitializeComponent();

            //抑制部分WebBrowerError提示視窗
            SuppressScriptErrorsOnly(webBrowser1);
            // Twitter API OAuth
            string consumerKey = "7CifoNPIwh4TPsHZ7HcLQ8BFL";
            string consumerSecret = "05496yBScmpgY6uI8i7GQp8QpVGWme1epWRzuWh2A0w4frpF7r";
            string token = "2902641404-BXzas8LsOTBr0SYWT2Fm1FqFqkshH0VeyXdnRd1";
            string tokenSecret = "yA062lZowtwPjFoi8NJ3un2nHsa4PkNLoZTJWJFhsEWxQ";
            servis = new TwitterService( consumerKey, consumerSecret, token, tokenSecret );
            
        }

        // 取得網頁原始碼
        private string GetHTMLSourceCode(string url)
        {
            HttpWebRequest request = (WebRequest.Create(url)) as HttpWebRequest;
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            using (StreamReader sr = new StreamReader(response.GetResponseStream()))
            {
                return sr.ReadToEnd();
            }
        }

        //載入瀏覽器
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl1.SelectedIndex == 0)
                {
                    textBox4.Text = dataGridView1.SelectedCells[0].Value.ToString();
                    string cellUrl = Convert.ToString(dataGridView1.SelectedCells[0].Value);
                    webBrowser1.Url = new Uri(cellUrl);
                }
                else if (tabControl1.SelectedIndex == 3)
                {
                    textBox4.Text = dataGridView2.SelectedCells[0].Value.ToString();
                    string cellUrl = Convert.ToString(dataGridView2.SelectedCells[0].Value);
                    webBrowser1.Url = new Uri(cellUrl);
                }    
            }
            catch(Exception err){
                
            }
        }
        private void SuppressScriptErrorsOnly(WebBrowser browser)
        {
            // Ensure that ScriptErrorsSuppressed is set to false.
            browser.ScriptErrorsSuppressed = false;

            // Handle DocumentCompleted to gain access to the Document object.
            browser.DocumentCompleted +=
                new WebBrowserDocumentCompletedEventHandler(
                    browser_DocumentCompleted);
        }

        private void browser_DocumentCompleted(object sender,
            WebBrowserDocumentCompletedEventArgs e)
        {
            ((WebBrowser)sender).Document.Window.Error +=
                new HtmlElementErrorEventHandler(Window_Error);
        }

        private void Window_Error(object sender,
            HtmlElementErrorEventArgs e)
        {
            // Ignore the error and suppress the error dialog box. 
            e.Handled = true;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            //初始化
            listBox1.Items.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 20;
            dataGridView1.RowCount = 200;
            String screen_name = textBox1.Text;
            int count = System.Convert.ToInt32( textBox2.Text );

            //取得資料
            var getTweets = servis.ListTweetsOnUserTimeline(
                new ListTweetsOnUserTimelineOptions() 
                { 
                    ScreenName = screen_name, Count = count, ExcludeReplies = true 
                });
            


            //欄位設定
            int no_num = 0;
            int url_link_num = 1;
            int idstr_num = 2;
            int author_num = 3;
            int text_num = 4;
            int entities_url_num = 5;
            int entities_tag_num = 6;
            dataGridView1.Columns[no_num].HeaderText = "NO.";
            dataGridView1.Columns[url_link_num].HeaderText = "推文連結";
            dataGridView1.Columns[idstr_num].HeaderText = "Idstr";
            dataGridView1.Columns[author_num].HeaderText = "Author";
            dataGridView1.Columns[text_num].HeaderText = "Text";
            dataGridView1.Columns[entities_url_num].HeaderText = "Entities_Url";
            dataGridView1.Columns[entities_tag_num].HeaderText = "Entities_Tag";


            int row_num = 0;
            int tweet_num = 0;
            foreach (var tweet in getTweets )
            {
                //No.
                dataGridView1.Rows[row_num].Cells[no_num].Value = tweet_num+1;

                //推文ID
                var idstr = tweet.IdStr;
                dataGridView1.Rows[row_num].Cells[idstr_num].Value = idstr;

                //推文作者
                var author = tweet.Author.ScreenName;
                dataGridView1.Rows[row_num].Cells[author_num].Value = author;

                //推文文字
                var text = tweet.Text;
                dataGridView1.Rows[row_num].Cells[text_num].Value = text;

                //推文(合併)連結
                var url_link = "http://twitter.com/" + author + "/status/" + idstr;
                dataGridView1.Rows[row_num].Cells[url_link_num].Value = url_link;

                int entitiesEndRow_num = row_num;
                int entitiesStartRow_num = row_num;
                var entities = tweet.Entities;
                
                //推文實體內文連結(外部連結)
                row_num = entitiesStartRow_num;
                var entities_urls = entities.Urls;
                foreach (var url in entities_urls)
                {
                    dataGridView1.Rows[row_num++].Cells[entities_url_num].Value = url.ExpandedValue;
                }
                row_num--;
                entitiesEndRow_num = entitiesEndRow_num>row_num?entitiesEndRow_num:row_num;

                //推文實體內文標籤
                row_num = entitiesStartRow_num;
                var entities_tags = entities.HashTags;
                foreach (var tag in entities_tags)
                {
                    dataGridView1.Rows[row_num++].Cells[entities_tag_num].Value = tag.Text;
                }
                row_num--;
                entitiesEndRow_num = entitiesEndRow_num > row_num ? entitiesEndRow_num : row_num;
                


                row_num = entitiesEndRow_num;
                row_num++;
                tweet_num++;

                //擴增列數
                if (row_num+20 >= dataGridView1.RowCount)
                {
                    dataGridView1.RowCount += 200;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
            dataGridView3.ColumnCount = 20;
            dataGridView3.RowCount = 200;
            dataGridView3.Columns[0].HeaderText = "收藏者";
            dataGridView3.Columns[1].HeaderText = "回覆者";
            dataGridView3.Columns[2].HeaderText = "回覆文字";
            HtmlDocument doc = webBrowser1.Document;
            //int a = doc.All.Count;
            string htmlfinder = "activity-popup-users dropdown-threshold";
            bool favoriteFound = false;
            for (int i = 0; i < doc.GetElementsByTagName("ol").Count; i++)
            {
                string htmlvalue = doc.GetElementsByTagName("ol")[i].OuterHtml;
                if (htmlvalue.Contains(htmlfinder))
                {
                    for (int j = 0, k = 0; j < doc.GetElementsByTagName("ol")[i].GetElementsByTagName("li").Count; j++)
                    {
                        if(doc.GetElementsByTagName("ol")[i].GetElementsByTagName("li")[j].GetElementsByTagName("div").Count > 0)
                        {
                            string screenName = doc.GetElementsByTagName("ol")[i].GetElementsByTagName("li")[j].GetElementsByTagName("div")[0].GetAttribute("data-screen-name");
                            dataGridView3.Rows[k++].Cells[0].Value = screenName;
                        }
                        //int hrefindex = htmlvalue.IndexOf("data-screen-name");
                        //for (int h = hrefindex + 18; htmlvalue[h] != '\"'; h++)
                        //{
                            //screenName += htmlvalue[h];
                        //}
                        
                    }
                    favoriteFound = true;
                }
                
            }
            if (favoriteFound) { }
            else
            {
                htmlfinder = "js-profile-popup-actionable";
                for(int i=0; i<doc.GetElementsByTagName("li").Count; i++)
                {
                    for(int j=0,k=0; j<doc.GetElementsByTagName("li")[i].GetElementsByTagName("a").Count; j++)
                    {
                        string htmlvalue = doc.GetElementsByTagName("li")[i].GetElementsByTagName("a")[j].OuterHtml;
                        if(htmlvalue.Contains(htmlfinder))
                        {
                            int hrefindex = htmlvalue.IndexOf("href");
                            string screenName = "";
                            for (int h = hrefindex + 7; htmlvalue[h] != '\"'; h++)
                            {
                                screenName += htmlvalue[h];
                            }
                            dataGridView3.Rows[k++].Cells[0].Value = screenName;
                        }
                    }
                
                }
            }
            
            htmlfinder = "js-simple-tweet";
            for (int i = 0,k=0; i < doc.GetElementsByTagName("li").Count; i++)
            {
                string htmlvalue = doc.GetElementsByTagName("li")[i].OuterHtml;
                if (htmlvalue.Contains(htmlfinder))
                {
                    htmlvalue = doc.GetElementsByTagName("li")[i].GetElementsByTagName("div")[0].OuterHtml;
                    string content = doc.GetElementsByTagName("li")[i].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].GetElementsByTagName("p")[0].OuterText;
                    int hrefindex = htmlvalue.IndexOf("screen-name");
                    string screenName = "";
                    for (int h = hrefindex + 13; htmlvalue[h] != '\"'; h++)
                    {
                        screenName += htmlvalue[h];
                    }
                    dataGridView3.Rows[k].Cells[1].Value = screenName;
                    dataGridView3.Rows[k].Cells[2].Value = content;
                    k++;
                }
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //初始化
            listBox1.Items.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.ColumnCount = 20;
            dataGridView2.RowCount = 200;
            var screen_name = textBox1.Text;
            var count = System.Convert.ToInt32(textBox2.Text);


            //取得資料
            var options = new ListFavoriteTweetsOptions()
            {
                ScreenName = screen_name,
                Count = count
            };
            var result = servis.BeginListFavoriteTweets(options);
            var getTweets = servis.EndListFavoriteTweets(result);
            /*var getTweets = servis.ListFavoriteTweets(
                new ListFavoriteTweetsOptions() 
                {
                    ScreenName = screen_name, Count = count
                });*/


            //欄位設定
            int no_num = 0;
            int url_link_num = 1;
            int idstr_num = 2;
            int author_num = 3;
            int text_num = 4;
            int entities_url_num = 5;
            int entities_tag_num = 6;
            dataGridView2.Columns[no_num].HeaderText = "NO.";
            dataGridView2.Columns[url_link_num].HeaderText = "推文連結";
            dataGridView2.Columns[idstr_num].HeaderText = "Idstr";
            dataGridView2.Columns[author_num].HeaderText = "Author";
            dataGridView2.Columns[text_num].HeaderText = "Text";
            dataGridView2.Columns[entities_url_num].HeaderText = "Entities_Url";
            dataGridView2.Columns[entities_tag_num].HeaderText = "Entities_Tag";


            int row_num = 0;
            int tweet_num = 0;
            foreach (var tweet in getTweets)
            {
                
                //No.
                dataGridView2.Rows[row_num].Cells[no_num].Value = tweet_num + 1;

                //推文ID
                var idstr = tweet.IdStr;
                dataGridView2.Rows[row_num].Cells[idstr_num].Value = idstr;

                //推文作者
                var author = tweet.Author.ScreenName;
                dataGridView2.Rows[row_num].Cells[author_num].Value = author;

                //推文文字
                var text = tweet.Text;
                dataGridView2.Rows[row_num].Cells[text_num].Value = text;

                //推文(合併)連結
                var url_link = "http://twitter.com/" + author + "/status/" + idstr;
                dataGridView2.Rows[row_num].Cells[url_link_num].Value = url_link;

                int entitiesEndRow_num = row_num;
                int entitiesStartRow_num = row_num;
                var entities = tweet.Entities;

                //推文實體內文連結(外部連結)
                row_num = entitiesStartRow_num;
                var entities_urls = entities.Urls;
                foreach (var url in entities_urls)
                {
                    dataGridView2.Rows[row_num++].Cells[entities_url_num].Value = url.ExpandedValue;
                }
                row_num--;
                entitiesEndRow_num = entitiesEndRow_num > row_num ? entitiesEndRow_num : row_num;

                //推文實體內文標籤
                row_num = entitiesStartRow_num;
                var entities_tags = entities.HashTags;
                foreach (var tag in entities_tags)
                {
                    dataGridView2.Rows[row_num++].Cells[entities_tag_num].Value = tag.Text;
                }
                row_num--;
                entitiesEndRow_num = entitiesEndRow_num > row_num ? entitiesEndRow_num : row_num;



                row_num = entitiesEndRow_num;
                row_num++;
                tweet_num++;

                //擴增列數
                if (row_num + 20 >= dataGridView2.RowCount)
                {
                    dataGridView2.RowCount += 200;
                }
            }
        }

        private void 儲存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    {
                        if (dataGridView1.Rows.Count == 0)
                        {
                            MessageBox.Show("No data available!", "Prompt", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                            saveFileDialog.FilterIndex = 0;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.CreatePrompt = true;
                            saveFileDialog.FileName = null;
                            saveFileDialog.Title = "Save path of the file to be exported";
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                Stream myStream = saveFileDialog.OpenFile();
                                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                                string strLine = "";
                                try
                                {
                                    //Write in the headers of the columns.  
                                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                                    {
                                        if (i > 0)
                                            strLine += ",";
                                        strLine += dataGridView1.Columns[i].HeaderText;
                                    }
                                    strLine.Remove(strLine.Length - 1);
                                    sw.WriteLine(strLine);
                                    strLine = "";
                                    //Write in the content of the columns.  
                                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                                    {
                                        strLine = "";
                                        for (int k = 0; k < dataGridView1.Columns.Count; k++)
                                        {
                                            if (k > 0)
                                                strLine += ",";
                                            if (dataGridView1.Rows[j].Cells[k].Value == null)
                                                strLine += "";
                                            else
                                            {
                                                string m = dataGridView1.Rows[j].Cells[k].Value.ToString().Trim();
                                                strLine += m.Replace(",", "，");
                                            }
                                        }
                                        strLine.Remove(strLine.Length - 1);
                                        sw.WriteLine(strLine);
                                        //Update the Progess Bar.  
                                        progressBar1.Value = 100 * (j + 1) / dataGridView1.Rows.Count;
                                    }
                                    sw.Close();
                                    myStream.Close();
                                    MessageBox.Show("Data has been exported to：" + saveFileDialog.FileName.ToString(), "Exporting Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    progressBar1.Value = 0;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Exporting Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }
                    break;
                case 2:
                    {
                        if (dataGridView3.Rows.Count == 0)
                        {
                            MessageBox.Show("No data available!", "Prompt", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                            saveFileDialog.FilterIndex = 0;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.CreatePrompt = true;
                            saveFileDialog.FileName = null;
                            saveFileDialog.Title = "Save path of the file to be exported";
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                Stream myStream = saveFileDialog.OpenFile();
                                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                                string strLine = "";
                                try
                                {
                                    //Write in the headers of the columns.  
                                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                                    {
                                        if (i > 0)
                                            strLine += ",";
                                        strLine += dataGridView3.Columns[i].HeaderText;
                                    }
                                    strLine.Remove(strLine.Length - 1);
                                    sw.WriteLine(strLine);
                                    strLine = "";
                                    //Write in the content of the columns.  
                                    for (int j = 0; j < dataGridView3.Rows.Count; j++)
                                    {
                                        strLine = "";
                                        for (int k = 0; k < dataGridView3.Columns.Count; k++)
                                        {
                                            if (k > 0)
                                                strLine += ",";
                                            if (dataGridView3.Rows[j].Cells[k].Value == null)
                                                strLine += "";
                                            else
                                            {
                                                string m = dataGridView3.Rows[j].Cells[k].Value.ToString().Trim();
                                                strLine += m.Replace(",", "，");
                                            }
                                        }
                                        strLine.Remove(strLine.Length - 1);
                                        sw.WriteLine(strLine);
                                        //Update the Progess Bar.  
                                        progressBar1.Value = 100 * (j + 1) / dataGridView3.Rows.Count;
                                    }
                                    sw.Close();
                                    myStream.Close();
                                    MessageBox.Show("Data has been exported to：" + saveFileDialog.FileName.ToString(), "Exporting Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    progressBar1.Value = 0;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Exporting Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }
                    break;
                case 3:
                    {
                        if (dataGridView2.Rows.Count == 0)
                        {
                            MessageBox.Show("No data available!", "Prompt", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                            saveFileDialog.FilterIndex = 0;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.CreatePrompt = true;
                            saveFileDialog.FileName = null;
                            saveFileDialog.Title = "Save path of the file to be exported";
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                Stream myStream = saveFileDialog.OpenFile();
                                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                                string strLine = "";
                                try
                                {
                                    //Write in the headers of the columns.  
                                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                                    {
                                        if (i > 0)
                                            strLine += ",";
                                        strLine += dataGridView2.Columns[i].HeaderText;
                                    }
                                    strLine.Remove(strLine.Length - 1);
                                    sw.WriteLine(strLine);
                                    strLine = "";
                                    //Write in the content of the columns.  
                                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                                    {
                                        strLine = "";
                                        for (int k = 0; k < dataGridView2.Columns.Count; k++)
                                        {
                                            if (k > 0)
                                                strLine += ",";
                                            if (dataGridView2.Rows[j].Cells[k].Value == null)
                                                strLine += "";
                                            else
                                            {
                                                string m = dataGridView2.Rows[j].Cells[k].Value.ToString().Trim();
                                                strLine += m.Replace(",", "，");
                                            }
                                        }
                                        strLine.Remove(strLine.Length - 1);
                                        sw.WriteLine(strLine);
                                        //Update the Progess Bar.  
                                        progressBar1.Value = 100 * (j + 1) / dataGridView2.Rows.Count;
                                    }
                                    sw.Close();
                                    myStream.Close();
                                    MessageBox.Show("Data has been exported to：" + saveFileDialog.FileName.ToString(), "Exporting Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    progressBar1.Value = 0;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Exporting Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }
                    break;
            }
              

        }

        private void 開啟ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileCSV = new OpenFileDialog();
            fileCSV.Title = "Open Excel File";
            if (fileCSV.ShowDialog() == DialogResult.OK)
            {
                string filename = System.IO.Path.GetFileName(fileCSV.FileName);
                string path = System.IO.Path.GetDirectoryName(fileCSV.FileName);
                
                
                //fileCSV.
                //Stream stream = 
                StreamReader reader = new StreamReader(path + "\\" + filename, System.Text.Encoding.GetEncoding(-0));
                
                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 20;
                dataGridView5.RowCount = 200;
                string line = reader.ReadLine();
                for (int j = 0, k = 0; j < line.Length && k < 20; j++, k++)
                {
                    string lineValue = "";
                    while (j < line.Length && line[j] != ',')
                    {
                        lineValue += line[j++];
                    }
                    if (lineValue != "")
                    {
                        dataGridView5.Columns[k].HeaderText = lineValue;
                    }
                }
                line = reader.ReadLine();
                int row_num = 0;
                while (line != null)
                {
                    for (int j = 0, k = 0; j < line.Length && k < 20; j++, k++)
                    {
                        string lineValue = "";
                        while (j < line.Length && line[j] != ',')
                        {
                            lineValue += line[j++];
                        }
                        if (lineValue != "")
                        {
                            dataGridView5.Rows[row_num].Cells[k].Value = lineValue;
                        }
                    }
                    line = reader.ReadLine();
                    row_num++;
                    //擴增列數
                    if (row_num + 20 >= dataGridView5.RowCount)
                    {
                        dataGridView5.RowCount += 200;
                    }
                }
                
            }
           /*OpenFileDialog fileDLG = new OpenFileDialog();  
           fileDLG.Title = "Open Excel File";  
           fileDLG.Filter = "Excel Files|*.xls;*.xlsx";  
           fileDLG.InitialDirectory = @"C:\Users\...\Desktop\";  
           if (fileDLG.ShowDialog() == DialogResult.OK)  
           {  
               string filename = System.IO.Path.GetFileName(fileDLG.FileName);  
               string path = System.IO.Path.GetDirectoryName(fileDLG.FileName);
               excelLocationTB.Text = @path + "\\" + filename;  
               string ExcelFile = @excelLocationTB.Text;  
               if (!File.Exists(ExcelFile))  
                   MessageBox.Show(String.Format("File {0} does not Exist", ExcelFile));

               System.Data.Odbc.OdbcConnection theConnection = new System.Data.Odbc.OdbcConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFile + ";Extended Properties=Excel 12.0;");  
               theConnection.Open();
               System.Data.Odbc.OdbcDataAdapter theDataAdapter = new System.Data.Odbc.OdbcDataAdapter("SELECT * FROM [Sheet1$]", theConnection);  
               DataSet DS = new DataSet();  
               theDataAdapter.Fill(DS, "ExcelInfo");  
               dataGridView1.DataSource = DS.Tables["ExcelInfo"];  
               formatDataGrid();  
               MessageBox.Show("Excel File Loaded");  
               progressBar1.Value += 0;  
           } */ 
       }

        private void button5_Click(object sender, EventArgs e)
        {
            HtmlDocument doc = webBrowser1.Document;
            string url = doc.Url.AbsoluteUri; 
            HttpWebRequest request = (WebRequest.Create(url)) as HttpWebRequest;
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            Stream Source = response.GetResponseStream();
            using (StreamReader sr = new StreamReader(Source,Encoding.GetEncoding(-0)))
            {
                string text = sr.ReadToEnd();
                richTextBox1.Text = text;
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            dataGridView4.Rows.Clear();
            var screen_name = textBox1.Text;
            var count = System.Convert.ToInt32(textBox2.Text);
            var gelen = servis.ListFollowers(new ListFollowersOptions() {ScreenName = screen_name });
            dataGridView4.ColumnCount = 20;
            dataGridView4.RowCount = 200;
            int i = 0;
            int followers_num = 0;
            dataGridView4.Columns[followers_num].HeaderText = "追隨者";
            
            foreach (var tweet in gelen)
            {

                var followers = tweet.ScreenName;
                dataGridView4.Rows[i].Cells[followers_num].Value = followers;
                
                i++;
            }
        }

        private void panel3_SizeChanged(object sender, EventArgs e)
        {
            int Wid = webBrowser1.Width;
            int Hig = panel3.Height - textBox4.Height;
            webBrowser1.Size = new System.Drawing.Size(Wid, Hig);
        }

        async private void splitter1_DoubleClick(object sender, EventArgs e)
        {
            int Wid = tabControl1.Width;
            int Hig = tabControl1.Height;
            int change_speed = 1;
            int path = Wid == 350 ? Wid - (panel2.Width - 350) : Wid - 350;
            
            for (int pixel = path ; pixel != 0; pixel=pixel>0?pixel-change_speed:pixel+change_speed)
            {
                tabControl1.Size = new System.Drawing.Size( (Wid-path) + pixel, Hig);
                
                if (pixel > 100 || pixel < -100)
                    change_speed = 4;
                else if (pixel < 50 || pixel > -50)
                    change_speed = 1;
                else
                    change_speed = 2;
                
            }
            tabControl1.Size = new System.Drawing.Size((Wid - path) , Hig);
        }

         
            /* private void formatDataGrid()  
            {  
             dataGridView1.ColumnHeadersVisible = true;  
             dataGridView1.Columns[0].Name = "Path Name";  
             dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);  

            }*/
        
    }
}
