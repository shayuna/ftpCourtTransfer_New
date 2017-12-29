using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop;

namespace ftpCourtTransfer_New
{
    public partial class Form1 : Form
    {
        private string[] arFolderNames;
        public Form1()
        {
            InitializeComponent();

            copyTo.Text = "c:\\users\\shay";
            DateTime dtYesterday = DateTime.Now.AddDays(-1);
            string sYesterday = dtYesterday.Day+"/"+ dtYesterday.Month+"/"+ dtYesterday.Year;
            fromDate.Text = sYesterday;
            toDate.Text = sYesterday;

            webBrowser1.Navigate ("https://decisions.court.gov.il");
            
        }
        private void DownloadFile(string sLocalFilePath, string sRemoteFileNm)
        {
            //            string inputfilepath = @"C:\tmp1234.docx";
            //          string httpsfilepath = "/2017-12-23//0012517460010000090037f69315425b.docx";
            
            string sHttpsHost = "decisions.court.gov.il";

            string sRemoteFilePath= "https://" + sHttpsHost + sRemoteFileNm;

            using (WebClient request = new WebClient())
            {
                request.Credentials = new NetworkCredential("idc\\data-hok", "L1150508a");
                byte[] fileData = request.DownloadData(sRemoteFilePath);
                using (FileStream file = File.Create(sLocalFilePath))
                {
                    file.Write(fileData, 0, fileData.Length);
                    file.Close();
                }
            }
        }
        private bool chkFldsValidity()
        {
            bool bValid = true;
            if (checkBox1.Checked && !IsDate(startAt.Text))
            {
                MessageBox.Show("עליך להזין תאריך ושעת התחלה לפי הפורמט המודגם");
                startAt.Focus();
                bValid = false;
            }
            else if (copyTo.Text == "")
            {
                MessageBox.Show("השדה העתקה לספריה ריק");
                copyTo.Focus();
                bValid = false;
            }
            else if (!Directory.Exists(copyTo.Text))
            {
                MessageBox.Show("ספריית היעד אינה קיימת");
                copyTo.Focus();
                bValid = false;
            }
            else if (fromDate.Text == "")
            {
                MessageBox.Show("שדה מתאריך ריק");
                copyTo.Focus();
                bValid = false;
            }
            else if (toDate.Text == "")
            {
                MessageBox.Show("שדה עד תאריך ריק");
                copyTo.Focus();
                bValid = false;
            }

            return bValid;
            /*
                ElseIf Trim(fromDate.Text) = "" Then
                    MsgBox ("äùãä 'îúàøéê' àéððå éëåì ìäéåú øé÷. òìéå ìäëéì àå îçøåæú áôåøîè úàøéê àå îçøåæú àçøú")
                    fromDate.SetFocus
                    chkFldsValidity = False
                ElseIf DateDiff("d", CDate(fromDate.Text), Date) < 1 And _
                        (Check1.Value = 0 Or DateDiff("d", CDate(fromDate.Text), CDate(startAt.Text)) < 1) Then
                    MsgBox ("òì äúàøéê áùãä 'îúàøéê' ìäéåú îå÷ãí îï äúàøéê ùì äéåí")
                    fromDate.SetFocus
                    chkFldsValidity = False
                ElseIf DateDiff("d", CDate(toDate.Text), Date) < 1 And _
                        (Check1.Value = 0 Or DateDiff("d", CDate(toDate.Text), CDate(startAt.Text)) < 1) Then
                    MsgBox ("òì äúàøéê áùãä 'òã úàøéê' ìäéåú îå÷ãí îï äúàøéê ùì äéåí")
                    toDate.SetFocus
                    chkFldsValidity = False
                ElseIf DateDiff("d", CDate(fromDate.Text), CDate(toDate.Text)) < 0 Then
                    MsgBox ("'äúàøéê áùãä 'òã úàøéê' îå÷ãí éåúø îï äúàøéê áùãä 'îúàøéê")
                    fromDate.SetFocus
                    chkFldsValidity = False
                End If



                         */





        }

        public static bool IsDate(Object obj)
        {
            string strDate = obj.ToString();
            try
            {
                DateTime dt = DateTime.ParseExact(strDate, "dd/MM/yyyy",CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label5.Text = "";
            if (chkFldsValidity())
            {
                bool bContinue = true;
                string copyToFolder = copyTo.Text;
                string sFolderNames = "";
                if (copyToFolder.Substring(copyToFolder.Length-1,1)!="/" && copyToFolder.Substring(copyToFolder.Length - 1, 1) != "\\")
                {
                    copyToFolder += "\\";
                }
                if (IsDate(fromDate.Text) && IsDate(toDate.Text))
                {
                    DateTime dtFromDate = DateTime.ParseExact(fromDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime dtToDate = DateTime.ParseExact(toDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    while (((TimeSpan)(dtToDate - dtFromDate)).Days >= 0)
                    {
                        string sFolderName = dtFromDate.Year + "-" + dtFromDate.Month + "-" + dtFromDate.Day;
                        if (sFolderNames != "") sFolderNames += ",";
                        sFolderNames += sFolderName;
                        sFolderName = copyToFolder + sFolderName;
                        dtFromDate=dtFromDate.AddDays(1);
                        Directory.CreateDirectory(sFolderName);
                    }
                }
                else if (!IsDate(fromDate.Text) && fromDate.Text != "")
                {
                    string sFolderName = copyToFolder + fromDate.Text;
                    Directory.CreateDirectory(sFolderName);
                    sFolderNames = sFolderName;
                }
                else
                {
                    bContinue = false;
                }



                /*
                                ElseIf Not IsDate(fromDate.Text) And fromDate.Text <> "" Then
                                      tmpFolderName = copyToFolder + fromDate.Text
                            fs.CreateFolder(tmpFolderName)
                            DateStr = fromDate.Text
                        Else
                            toCont = False
                        End If
                */
                if (bContinue)
                {
                    label5.Text = "סטטוס: משיכת קבצים";
                    arFolderNames = sFolderNames.Split(',');
                    foreach (string sFolder in arFolderNames)
                    {
                        string tmpUrl = "https://decisions.court.gov.il/" + arFolderNames[0] + "/";
                        webBrowser1.Navigate(tmpUrl);
                        while (webBrowser1.IsBusy || webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                        {
                            Application.DoEvents();
                            Application.DoEvents();
                            Application.DoEvents();
                            //Thread.Sleep(1000);
                        }
                        downloadFiles(sFolder);
                    }
                    label5.Text = "";
                    MessageBox.Show("תהליך ההורדה הסתיים");
                }


            }
            else
            {
            }
        }
        private void downloadFiles(string sFolderNm)
        {
            string tmpTxt = webBrowser1.Document.Body.InnerHtml;
            int iCount = 0;
            Regex re = new Regex("<a.*?</a>",RegexOptions.IgnoreCase);
            Regex re2 = new Regex(@"href=""(.*?)""", RegexOptions.IgnoreCase);
            Regex re3 = new Regex(">(.*?)<", RegexOptions.IgnoreCase);
            MatchCollection mt = re.Matches(tmpTxt);
            foreach (Match m in mt)
            {
                string sLink = m.ToString();
                if (sLink.IndexOf("href") > -1 && (sLink.IndexOf("doc")>-1 || sLink.IndexOf("xml") > -1))
                {
                    iCount++;
                    label5.Text = "משיכת קובץ " + iCount + " מתוך " + mt.Count + " קבצים";
                    label5.Refresh();
                    string sHref = re2.Match(sLink).Groups[1].ToString();
                    string sName = re3.Match(sLink).Groups[1].ToString();
                    string sLocalPath = copyTo.Text + "\\"+sFolderNm+"\\"+sName;
                    DownloadFile(sLocalPath, sHref);
                    Thread.Sleep(1000);
                }
            }
            

            
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            label5.Text = "";
            bool bContinue = true;
            string sRemoteFolderNm="",sLocalFolderNm="",sLocalFolderPath="";

            if (IsDate(fromDate.Text) && IsDate(toDate.Text))
            {
                DateTime dtFromDate = DateTime.ParseExact(fromDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime dtToDate = DateTime.ParseExact(toDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (((TimeSpan)(dtToDate - dtFromDate)).Days != 0)
                {
                    bContinue = false;
                }
                else
                {
                    sRemoteFolderNm = dtFromDate.Year + "-" + dtFromDate.Month + "-" + dtFromDate.Day;
                    sLocalFolderNm = sRemoteFolderNm;
                }

            }
            else
            {
                sRemoteFolderNm = fromDate.Text;
                sLocalFolderNm = sRemoteFolderNm;
            }

            sLocalFolderPath = copyTo.Text + "//" + sLocalFolderNm+"//";

            if (!Directory.Exists(sLocalFolderPath))
            {
                bContinue=false;
            }
            else if (chkFldsValidity() && bContinue)
            {

                label5.Text = "סטטוס: השוואת ספריות";
                string tmpUrl = "https://decisions.court.gov.il/" + sRemoteFolderNm + "/";
                bool bAllFilesDownloaded = true;
                webBrowser1.Navigate(tmpUrl);
                while (webBrowser1.IsBusy || webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                {
                    Application.DoEvents();
                    Application.DoEvents();
                    Application.DoEvents();
                    //Thread.Sleep(1000);
                }


                string tmpTxt = webBrowser1.Document.Body.InnerHtml;
                List<string> arFiles = new List<string>(); 
                Regex re = new Regex("<a.*?</a>", RegexOptions.IgnoreCase);
                Regex re2 = new Regex(@"href=""(.*?)""", RegexOptions.IgnoreCase);
                Regex re3 = new Regex(">(.*?)<", RegexOptions.IgnoreCase);
                MatchCollection mt = re.Matches(tmpTxt);

                foreach (Match m in mt)
                {
                    string sLink = m.ToString();
                    if (sLink.IndexOf("href") > -1 && (sLink.IndexOf("doc") > -1 || sLink.IndexOf("xml") > -1))
                    {
                        string sHref = re2.Match(sLink).Groups[1].ToString();
                        string sName = re3.Match(sLink).Groups[1].ToString();
                        arFiles.Add(sHref);
                    }
                }

                foreach (string sFile in arFiles)
                {
                    if (!File.Exists(copyTo.Text + "//" + sFile))
                    {
                        bAllFilesDownloaded = false;
                        DownloadFile(copyTo.Text + "//"+sFile, sFile);
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        Console.WriteLine(sFile + " exists");
                    }
                }
                if (bAllFilesDownloaded)
                {
                    MessageBox.Show("שתי הספריות נמצאו זהות");
                }
                else
                {
                    MessageBox.Show("תהליך ההשוואה והשלמת החסרים הסתיים בהצלחה");
                }
            }
            if (!bContinue)
            {
                MessageBox.Show("תהליך ההשוואה נכשל בשל אחת משתי סיבות אפשריות: הספריה המקומית אינה קיימת או שטווח התאריכים שנבחר גדול מיום אחד");
            }
            label5.Text = "";

        }
    }


}


