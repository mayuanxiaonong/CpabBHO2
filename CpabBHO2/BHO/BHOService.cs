using System;
using System.Collections.Generic;
using System.Text;

using SHDocVw;
using mshtml;

namespace CpabBHO2.BHO
{
    /* BHO插件的逻辑处理 */
    public class BHOService : IExtension
    {
        public WebBrowser webBrowser;
        public HTMLDocument document;

        /* 注入JS脚本 */
        public void registerExtensionJS()
        {
            document = (HTMLDocument)webBrowser.Document;
            
            
            document.parentWindow.execScript(
                            @"$(""form#Indvcustomer1 button[type='submit']"").click(function(){
                        window.cpabExtension.SaveCustomer("""");
                    });", "JavaScript");

            document.parentWindow.execScript(
                            @"$(""form#Indvcustomer1 #tbCsmCustomer_certificateCode"").blur(function(){
alert(""Blur "" + this.value);
                        window.cpabExtension.BlurCusNo("""");
                    });", "JavaScript");


            document.parentWindow.execScript(
                            @"$(""div.loginFormBody #j_username"").blur(function(){
                        window.cpabExtension.BlurCusNo("""");
                    });", "JavaScript");

            //                }
            //            }

            document.parentWindow.execScript(
                @"$(document).on(""click"", ""#appendbtn"", function(){
			        window.cpabExtension.Foo(this.id);
                });", "JavaScript");
            
            //document.parentWindow.execScript(this.Js_Mobile);
        }

        public void Foo(string s)
        {
            System.Windows.Forms.MessageBox.Show("BHO " + s);
        }

        public void SaveCustomer(string s)
        {
            
            //System.Windows.Forms.MessageBox.Show("已拦截到按钮的点击操作");
            
            document = (HTMLDocument)webBrowser.Document;

            /*
ECIF客户编号   tbCsmCustomer_cusNo  input
客户名称  tbCsmCustomer_cusName   input
证件类型  tbCsmCustomer_certificateType   select
证件号码  tbCsmCustomer_certificateCode   input
性别  tbCsmIndvCustomer_gender  select
出生日期  tbCsmIndvCustomer_birthday  input   yyyy-MM-dd
户籍所在地  tbCsmIndvCustomer_nativePlace   input
            */
            string cusNo = "", cusName="", certType="", certTypeStr="", certCode="", sex="", sexStr="", birthday="", nativePlace="";

            try
            {
                // ECIF客户编号
                IHTMLInputElement eCusNo = (IHTMLInputElement)document.getElementById("tbCsmCustomer_cusNo");
                cusNo = eCusNo == null ? "" : eCusNo.value;
                // 客户名称
                IHTMLInputElement eCusName = (IHTMLInputElement)document.getElementById("tbCsmCustomer_cusName");
                cusName = eCusName == null ? "" : eCusName.value;
                // 证件类型
                IHTMLSelectElement eCertType = (IHTMLSelectElement)document.getElementById("tbCsmCustomer_certificateType");
                certType = eCertType == null ? "" : eCertType.value;
                certTypeStr = cerType.ContainsKey(certType) ? cerType[certType] : "";
                // 证件号码
                IHTMLInputElement eCertCode = (IHTMLInputElement)document.getElementById("tbCsmCustomer_certificateCode");
                certCode = eCertCode == null ? "" : eCertCode.value;
                // 性别
                IHTMLSelectElement eSex = (IHTMLSelectElement)document.getElementById("tbCsmIndvCustomer_gender");
                sex = eSex == null ? "" : eSex.value;
                sexStr = sexType.ContainsKey(sex) ? sexType[sex] : "";
                // 出生日期
                IHTMLInputElement eBirth = (IHTMLInputElement)document.getElementById("tbCsmIndvCustomer_birthday");
                birthday = eBirth == null ? "" : eBirth.value;
                // 户籍所在地
                IHTMLInputElement eNativePlace = (IHTMLInputElement)document.getElementById("tbCsmIndvCustomer_nativePlace");
                nativePlace = eNativePlace == null ? "" : eNativePlace.value;
            }
            catch (Exception e) {
                System.Windows.Forms.MessageBox.Show("BHO - Save" + e.ToString());
            }

            string info = string.Format("ECIF客户编号: {0}\r\n客户名称: {1}\r\n证件类型: {2} {3}\r\n证件号码: {4}\r\n性别: {5} {6}\r\n出生日期: {7}\r\n户籍所在地: {8}\r\n",
                cusNo, cusName, certType, certTypeStr, certCode, sex, sexStr, birthday, nativePlace);
            
            System.Windows.Forms.MessageBox.Show("[Message from BHO] - 用户基本信息：\r\n" + info);
        }

        public void BlurCusNo(string s)
        {
            try
            {
                // 证件号码
                IHTMLInputElement eCertCode = (IHTMLInputElement)document.getElementById("tbCsmCustomer_certificateCode");
                string certCode = eCertCode == null ? "" : eCertCode.value;
                System.Windows.Forms.MessageBox.Show("BHO，证件号码：" + certCode);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("BHO-Blur "+e.ToString());
            }
        }

        public string Js_Mobile = @"$(""#mobile"").blur(function() {
                    window.myExtension.Mobile(this.value);
                });";
        public void Mobile(string s)
        {
            MobileThread mt = new MobileThread(s);
            new System.Threading.Thread(new System.Threading.ThreadStart(mt.http)).Start();
        }

        static Dictionary<string, string> cerType = new Dictionary<string, string>();
        static Dictionary<string, string> sexType = new Dictionary<string, string>();

        static BHOService(){
            cerType.Add("10", "居民身份证");
            cerType.Add("31", "武警士兵证");
            cerType.Add("32", "警官证");
            cerType.Add("33", "武警文职干部证");
            cerType.Add("34", "武警军官退休证");
            cerType.Add("35", "武警文职干部退休证");
            cerType.Add("40", "户口簿");
            cerType.Add("50", "本国护照");
            cerType.Add("51", "外国护照");
            cerType.Add("60", "港澳居民来往内地通行证");
            cerType.Add("61", "台湾居民来往大陆通行证");
            cerType.Add("11", "临时身份证");
            cerType.Add("70", "外国人永久居留证");
            cerType.Add("99", "其他");
            cerType.Add("20", "军人身份证");
            cerType.Add("21", "士兵证");
            cerType.Add("22", "军官证");
            cerType.Add("23", "文职干部证");
            cerType.Add("24", "军官退休证");
            cerType.Add("25", "文职干部退休证");
            cerType.Add("30", "武警身份证");

            sexType.Add("0", "未知的性别");
            sexType.Add("1", "男性");
            sexType.Add("2", "女性");
        }
    }
}
