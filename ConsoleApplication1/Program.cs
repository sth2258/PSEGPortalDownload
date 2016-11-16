using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using WatiN.Core;
using WatiN.Core.DialogHandlers;

namespace ConsoleApplication1
{
    class Program
    {

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        private static Dictionary<string, string> urls;
        private static string singleGetReportURL;

        [STAThread]
        static void Main(string[] args)
        {
            urls = new Dictionary<string, string>();
            urls.Add("2-Review", "https://www.pseglinyportal.com/Reserved.ReportViewerWebControl.axd?ReportSession=%ReportSession%&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=%ControlID%&OpType=Export&FileName=Rpt_ProjectDataExtract_Drilldown&ContentDisposition=OnlyHtmlInline&Format=CSV");
            urls.Add("4-Committed", "https://www.pseglinyportal.com/Reserved.ReportViewerWebControl.axd?ReportSession=%ReportSession%&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=%ControlID%&OpType=Export&FileName=Rpt_ProjectDataExtract_Drilldown&ContentDisposition=OnlyHtmlInline&Format=CSV");
            urls.Add("5-Installed", "https://www.pseglinyportal.com/Reserved.ReportViewerWebControl.axd?ReportSession=%ReportSession%&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=%ControlID%&OpType=Export&FileName=Rpt_ProjectDataExtract_Drilldown&ContentDisposition=OnlyHtmlInline&Format=CSV");
            urls.Add("OnHold", "https://www.pseglinyportal.com/Reserved.ReportViewerWebControl.axd?ReportSession=%ReportSession%&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=%ControlID%&OpType=Export&FileName=Rpt_ProjectDataExtract_Drilldown&ContentDisposition=OnlyHtmlInline&Format=CSV");

            singleGetReportURL = "https://www.pseglinyportal.com/Internal/DrillDownReport.aspx?Status=2-Review";

            DownloadFiles();


        }

        private static void DownloadFiles()
        {
            var processes = from process in System.Diagnostics.Process.GetProcesses()
                            where process.ProcessName == "iexplore"
                            select process;

            foreach (var process in processes)
            {
                while (!process.HasExited)
                {
                    process.Kill();
                    process.WaitForExit();
                }
            }

            var browser = new IE("https://www.pseglinyportal.com/Account/Login.aspx");

            browser.TextField(Find.ById("MainContent_UserName")).TypeText("ext-habera");
            browser.TextField(Find.ById("MainContent_Password")).TypeText("917-295-1167pseg");

            //browser.Button(Find.ByName("MainContent_btnsubmit")).Click();
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            Console.WriteLine("Enter submitted");
            Thread.Sleep(30000);
            //browser.WaitForComplete();

            //for the first URL in the list, we need to download the html of the root report viewer
            //from there, we extract the "ReportSession" ID, then use that to pass to the report server so we can dynamically download the exact report we want

            browser.GoTo(singleGetReportURL);
            Console.WriteLine("Waiting to load " + singleGetReportURL);
            Thread.Sleep(30000);
            WriteFile(browser.Html, "tmp.html");

            string reportSessionID = GetTargetValue("tmp.html", "ReportSession=");
            string controlID = GetTargetValue("tmp.html", "ControlID=");

            //noenw that we've got the session ID, execute the proper GET requests to download our CSV's 

            foreach (var entry in urls)
            {
                string url = entry.Value.Replace("%ReportSession%", reportSessionID);
                url = url.Replace("%ControlID%", controlID);


               // DownLoadFile(ref browser, url);

                browser.GoTo(url);
                DownLoadFile("Drill Down Report - Internet Explorer");

                Console.WriteLine("Waiting to load " + entry.Key);
                Thread.Sleep(30000);
                //WriteFile(browser.Html, entry.Key + ".html");
            }
            browser.Close();
        }

        private static void DownLoadFile(string strWindowTitle)
        {
            IntPtr TargetHandle = FindWindowByCaption(IntPtr.Zero, strWindowTitle);
            AutomationElementCollection ParentElements = AutomationElement.FromHandle(TargetHandle).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement ParentElement in ParentElements)
            {
                // Identidfy Download Manager Window in Internet Explorer
                if (ParentElement.Current.ClassName == "Frame Notification Bar")
                {
                    AutomationElementCollection ChildElements = ParentElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    // Idenfify child window with the name Notification Bar or class name as DirectUIHWND 
                    foreach (AutomationElement ChildElement in ChildElements)
                    {
                        if (ChildElement.Current.Name == "Notification bar" || ChildElement.Current.ClassName == "DirectUIHWND")
                        {

                            AutomationElementCollection DownloadCtrls = ChildElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                            foreach (AutomationElement ctrlButton in DownloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                if (ctrlButton.Current.Name.ToLower() == "save")
                                {
                                    var invokePattern = ctrlButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                    invokePattern.Invoke();
                                }

                            }
                        }
                    }


                }
            }
        }

        static void DownLoadFile(ref IE browser, string uri)
        {
            browser.GoTo(uri);

            Thread.Sleep(1000);
            AutomationElementCollection dialogElements = AutomationElement.FromHandle(FindWindow(null, "Drill Down Report - Internet Explorer")).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement element in dialogElements)
            {
                if (element.Current.Name.Equals("Save"))
                { 
                    var invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern.Invoke();

                }
            }
        }

        private static void WriteFile(string body, string fileName)
        {
            StreamWriter wrt = new StreamWriter(fileName);
            wrt.WriteLine(body);
            wrt.Flush();
            wrt.Close();
        }

        private static string GetTargetValue(string htmlfile, string target)
        {
            StreamReader rdr = new StreamReader(htmlfile);
            string fullText = rdr.ReadToEnd();
            int startIndex = fullText.IndexOf(target) + target.Length;
            int endIndex = fullText.IndexOf("\\", startIndex + 1);
            string reportSession = fullText.Substring(startIndex, endIndex - startIndex);

            return reportSession;
        }
    }
}
