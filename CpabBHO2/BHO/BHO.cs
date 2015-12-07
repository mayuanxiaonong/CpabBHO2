using System;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Expando;
using System.Reflection;
using System.Net;
using System.Text.RegularExpressions;

using SHDocVw;
using mshtml;

namespace CpabBHO2.BHO
{

    [
        ComVisible(true),
        Guid("34e09a14-2cda-4334-8f02-48f450fc0560"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("MyExtension"),
        ComDefaultInterface(typeof(IExtension))
        ]

    public class BHO : BHOService,IObjectWithSite
    {

        public static string _URL = "";
        public bool IS_INJECT = false;
        
        // 注册外部js函数
        public void registerExtensionService()
        {
            
            try
            {
                IHTMLWindow2 window = ((IHTMLDocument2)webBrowser.Document).parentWindow;
                IExpando windowEx = (IExpando)window;
                PropertyInfo myExtension = windowEx.GetProperty("cpabExtension", BindingFlags.Default);
                if (myExtension == null)
                {
                    myExtension = windowEx.AddProperty("cpabExtension");
                }
                myExtension.SetValue(windowEx, this, null);
            }
            catch { }
        }

        /* 注册刷新事件 */
        public void regRefresh()
        {
            HTMLDocument doc = webBrowser.Document as HTMLDocument;
            if (doc != null)
            {
                IHTMLWindow2 tmpWindow = doc.parentWindow;
                if (tmpWindow != null)
                {
                    HTMLWindowEvents2_Event events = (tmpWindow as HTMLWindowEvents2_Event);
                    try
                    {
                        events.onload -= new HTMLWindowEvents2_onloadEventHandler(this.OnRefreshHandler);
                    }
                    catch { }
                    events.onload += new HTMLWindowEvents2_onloadEventHandler(this.OnRefreshHandler);
                }
            }
        }

        /* 注册扩展 */
        public void registerExtension()
        {
            //System.Windows.Forms.MessageBox.Show("IS_INJECT " + IS_INJECT);
            if(!IS_INJECT)
            {
                IS_INJECT = isUrlMatch(_URL);
            }
            //if (IS_INJECT)
            //{
            
                document = (HTMLDocument)webBrowser.Document;
                IHTMLElement e = document.getElementById("_BHO_");
                if (e == null)
                {
                    IHTMLElement body = document.body;
                    if (body != null)
                    {
                        try
                        {
                            this.registerExtensionService();
                            this.registerExtensionJS();
                            
                            body.insertAdjacentHTML("afterBegin", "<input id=\"_BHO_\" type=\"hidden\" value=\"" + DateTime.Now.ToString() + "\" />");
                        }
                        catch (Exception) {
                        System.Windows.Forms.MessageBox.Show("Err " + e.ToString());
                        }
                    }
                }


            //}
        }

        /* 判断url是否匹配注入规则 */
        public bool isUrlMatch(string url)
        {
            // 个贷系统地址
            Regex reg = new Regex("^http\\://.*21\\.6\\.11\\.6/easyloan/.*");
            //Regex reg = new Regex(".*/index.*");
            if (!reg.IsMatch(url))
            {
                //System.Windows.Forms.MessageBox.Show("BHO " + url);
                return false;
            }
            else
            {
                return true;
            }
        }


        public void OnDocumentComplete(object pDisp, ref object URL)
        {
        }
        public int i = 1;
        public void OnDownloadComplete()
        {
            try {
                this.registerExtension();
            }
            catch(Exception)
            { }
            //this.regRefresh();
        }

        public void OnDownloadBegin()
        {
        }

        public void OnBeforeNavigate2(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            _URL = URL as string;
        }

        public void OnRefreshHandler(IHTMLEventObj e)
        {
        }

        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
                webBrowser.DownloadBegin += new DWebBrowserEvents2_DownloadBeginEventHandler(this.OnDownloadBegin);
                webBrowser.DownloadComplete += new DWebBrowserEvents2_DownloadCompleteEventHandler(this.OnDownloadComplete);
            }
            else
            {
                webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser.BeforeNavigate2 -= new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
                webBrowser.DownloadBegin -= new DWebBrowserEvents2_DownloadBeginEventHandler(this.OnDownloadBegin);
                webBrowser.DownloadComplete -= new DWebBrowserEvents2_DownloadCompleteEventHandler(this.OnDownloadComplete);
                webBrowser = null;
            }
            return 0;
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }

        public static string BHO_KEY = @"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHO_KEY, true);
            if (registryKey == null)
            {
                registryKey = Registry.LocalMachine.CreateSubKey(BHO_KEY);
            }


            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid, true);
            if (ourKey == null)
            {
                ourKey = registryKey.CreateSubKey(guid);
            }

            ourKey.SetValue("NoExplorer", 1);
            //ourKey.SetValue("Alright", 1);

            registryKey.Close();
            ourKey.Close();
            

            RegistryKey settingKey = Registry.CurrentUser.OpenSubKey(@"Software\\Microsoft\\Windows\\CurrentVersion\\Ext\\Settings", true);
            if (settingKey != null)
            {
                RegistryKey rKey = settingKey.OpenSubKey("");
                if (rKey == null)
                {
                    rKey = settingKey.CreateSubKey(guid);
                }
                rKey.SetValue("Flags", 0x400);

                rKey.Close();
            }
            settingKey.Close();

        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHO_KEY, true);
            string guid = type.GUID.ToString("B");
            if (registryKey != null)
            {
                registryKey.DeleteSubKey(guid, false);
            }
        }

    }
}
