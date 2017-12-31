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
using Microsoft.Office.Interop.Word;

namespace ftpCourtTransfer_New
{
    public partial class Form1 : Form
    {
        private string[] arFolderNames;
        Microsoft.Office.Interop.Word.Application oWordApp;
        public Form1()
        {
            InitializeComponent();
            copyTo.Text = "";
            DateTime dtYesterday = DateTime.Now.AddDays(-1);
            DateTime dtNow = DateTime.Now;
            string sNow = dtNow.Day + "/" + dtNow.Month + "/" + dtNow.Year + " " + dtNow.Hour + ":" + dtNow.Minute + ":" + dtNow.Second;
            string sYesterday = dtYesterday.Day+"/"+ dtYesterday.Month+"/"+ dtYesterday.Year;
            fromDate.Text = sYesterday;
            toDate.Text = sYesterday;
            startAt.Text = sNow;
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
                    moveDocHdr(sLocalFilePath);
                }
            }
        }
        private bool chkFldsValidity()
        {
            bool bValid = true;
            DateTime dtFromDate=DateTime.Now.Date,dtToDate = DateTime.Now.Date, dtStartAt = DateTime.Now.Date;
            if (IsDate(fromDate.Text)) dtFromDate = DateTime.ParseExact(fromDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (IsDate(toDate.Text)) dtToDate = DateTime.ParseExact(toDate.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (IsDateWithTime(startAt.Text)) dtStartAt = DateTime.ParseExact(startAt.Text, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture).Date;

            if (checkBox1.Checked && !IsDateWithTime(startAt.Text))
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
            else if (!IsDate(fromDate.Text))
            {
                MessageBox.Show("על התאריך בשדה 'מתאריך' להיות לפי הפורמט המודגם");
                fromDate.Focus();
                bValid = false;
            }
            else if (!IsDate(toDate.Text))
            {
                MessageBox.Show("על התאריך בשדה 'עד תאריך' להיות לפי הפורמט המודגם");
                toDate.Focus();
                bValid = false;
            }
            else if (!checkBox1.Checked && DateTime.Compare(dtFromDate,DateTime.Now.Date)>=0)
            {
                MessageBox.Show("על התאריך בשדה 'מתאריך' להיות מוקדם מהתאריך הנוכחי");
                fromDate.Focus();
                bValid = false;
            }
            else if (!checkBox1.Checked  && DateTime.Compare(dtToDate, DateTime.Now.Date) >= 0)
            {
                MessageBox.Show("על התאריך בשדה 'עד תאריך' להיות מוקדם מהתאריך הנוכחי");
                toDate.Focus();
                bValid = false;
            }
            else if (checkBox1.Checked && DateTime.Compare(dtFromDate, dtStartAt) >= 0)
            {
                MessageBox.Show("על התאריך בשדה 'מתאריך' להיות מוקדם יותר משדה 'התחל ב");
                fromDate.Focus();
                bValid = false;
            }
            else if (checkBox1.Checked && DateTime.Compare(dtToDate, dtStartAt) >= 0)
            {
                MessageBox.Show("על התאריך בשדה 'עד תאריך' להיות מוקדם יותר משדה 'התחל ב");
                toDate.Focus();
                bValid = false;
            }
            else if (DateTime.Compare(dtFromDate, dtToDate) >0)
            {
                MessageBox.Show("על התאריך בשדה 'מתאריך' להיות מוקדם או שווה לשדה 'עד תאריך");
                toDate.Focus();
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
                DateTime dt = DateTime.ParseExact(strDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public static bool IsDateWithTime(Object obj)
        {
            string strDate = obj.ToString();
            try
            {
                DateTime dt = DateTime.ParseExact(strDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }
        private void startCopyingFiles()
        {
            bool bContinue = true;
            string copyToFolder = copyTo.Text;
            string sFolderNames = "";
            if (copyToFolder.Substring(copyToFolder.Length - 1, 1) != "/" && copyToFolder.Substring(copyToFolder.Length - 1, 1) != "\\")
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
                    dtFromDate = dtFromDate.AddDays(1);
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
                        System.Windows.Forms.Application.DoEvents();
                        System.Windows.Forms.Application.DoEvents();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    downloadFiles(sFolder);
                }
                label5.Text = "";
                MessageBox.Show("תהליך ההורדה הסתיים");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            /*
                        moveDocHdr("c:\\users\\shay\\tmpdoc.docx");
                        System.Windows.Forms.Application.Exit();
                        return;
              */

            if (!chkFldsValidity()) return;

            if (checkBox1.Checked)
            {
                DateTime dtStart = DateTime.ParseExact(startAt.Text, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                while (DateTime.Compare(dtStart, DateTime.Now) > 0)
                {
                    label5.Text = "מתחילים להוריד קבצים בעוד "+Math.Round(((TimeSpan)(dtStart-DateTime.Now)).TotalMinutes)+" דקות";
                    label5.Refresh();
                    //                    Thread.Sleep(10000);
                    new System.Threading.ManualResetEvent(false).WaitOne(10000);
                }
            }
            label5.Text = "הורדת הקבצים תחל כעת";
            label5.Refresh();

            webBrowser1.Navigate("https://decisions.court.gov.il");


            startCopyingFiles();
        }
        private void downloadFiles(string sFolderNm)
        {
            string tmpTxt = webBrowser1.Document.Body.InnerHtml;
            int iCount = 0;
            Regex re = new Regex("<a.*?</a>",RegexOptions.IgnoreCase);
            Regex re2 = new Regex(@"href=""(.*?)""", RegexOptions.IgnoreCase);
            Regex re3 = new Regex(">(.*?)<", RegexOptions.IgnoreCase);
            MatchCollection mt = re.Matches(tmpTxt);
            oWordApp = new Microsoft.Office.Interop.Word.Application();
            oWordApp.Visible = false;
            oWordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
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
                    //                    Thread.Sleep(1000);
                    new System.Threading.ManualResetEvent(false).WaitOne(1000);
                }
            }
            oWordApp.Quit(WdSaveOptions.wdSaveChanges);
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
                    System.Windows.Forms.Application.DoEvents();
                    System.Windows.Forms.Application.DoEvents();
                    System.Windows.Forms.Application.DoEvents();
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

                oWordApp = new Microsoft.Office.Interop.Word.Application();
                oWordApp.Visible = false;
                oWordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                foreach (string sFile in arFiles)
                {
                    if (!File.Exists(copyTo.Text + "//" + sFile))
                    {
                        bAllFilesDownloaded = false;
                        DownloadFile(copyTo.Text + "//"+sFile, sFile);
                        //                    Thread.Sleep(1000);
                        new System.Threading.ManualResetEvent(false).WaitOne(1000);
                    }
                    else
                    {
                        Console.WriteLine(sFile + " exists");
                    }
                }
                oWordApp.Quit(WdSaveOptions.wdSaveChanges);
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
        private void moveDocHdr(string sLocalFilePath)
        {
            Microsoft.Office.Interop.Word.Document oDoc = oWordApp.Documents.Open(sLocalFilePath);
            try
            {
                Regex re = new Regex("[a-z0-9א-ת]", RegexOptions.IgnoreCase);
                bool bEditeHdrSuccess = true;
                if (oDoc.Sections[1].Headers.Count > 0 && re.IsMatch(oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text))
                {
                    Range oRng = oDoc.Range(Start: 0, End: 0);
                    int iPrgBefore = oDoc.Paragraphs.Count, iPrgNum = 0, jj = 0, iPrgAfter = 0;
                    try
                    {
                        if (oRng.Tables.Count == 1)
                        {
                            oRng.InsertBreak(Type: WdBreakType.wdColumnBreak);
                        }
                        else
                        {
                            oRng.InsertBreak(Type: WdBreakType.wdLineBreak);
                        }
                    }
                    catch (Exception ex)
                    {
                        //                    MessageBox.Show(sLocalFilePath);
                    }
                    iPrgNum = oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Count;
                    Clipboard.Clear();
                    try
                    {
                        oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Cut();
                        oRng.Paste();
                    }
                    catch (Exception ex)
                    {
                        bEditeHdrSuccess = false;
                        /* in case editing isn't allowed on this part of the document */
                    }
                    if (!bEditeHdrSuccess)
                    {
                        oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Copy();
                        oRng.Paste();
                        oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
                    }
                    iPrgAfter = oDoc.Paragraphs.Count;
                    for (jj = 1; jj <= iPrgAfter - iPrgBefore; jj++)
                    {
                        oDoc.Paragraphs[jj].NoLineNumber = -1;
                    }
                    Clipboard.Clear();
                }
                oDoc.Close(WdSaveOptions.wdSaveChanges);
            }
            catch (Exception ex)
            {
                writeToLog(sLocalFilePath);
                oDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                if (File.Exists(sLocalFilePath)) File.Delete(sLocalFilePath);
                //general exception 
            }
        }
        private void writeToLog(string sTxt)
        {
            string sLogFlNm = System.IO.Directory.GetCurrentDirectory() + "//" + "ftpCourtTransfer_New_log.txt";
            StreamWriter flLog = null;
            if (!File.Exists(sLogFlNm)) flLog = File.CreateText(sLogFlNm);
            else flLog = File.AppendText(sLogFlNm);
            flLog.WriteLine(sTxt);
            flLog.Close();
        }



    }


}


