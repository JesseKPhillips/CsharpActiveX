using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ActiveXObjectSpace;

namespace TestActiveX
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class Form1 : Form
    {
        MyObject myObject = new MyObject();
        public Form1()
        {
            InitializeComponent();
            Text = "ActiveX Test";

            Load += Form1_Load;
        }

        /*
         * The ActiveX objects will not fire the javascript event handler
         * So we will insert javascript functions which will grab the
         * Javascript object for us.
         */
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {
                myObject = this.webBrowser1.Document.InvokeScript("eval", new[] { "MyObject.object" }) as MyObject;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.AllowWebBrowserDrop = false;
            webBrowser1.ObjectForScripting = this;
            webBrowser1.Url = new Uri(@"file:///c|\tmp\CsharpActiveX\TestActiveX\TestPage.html");

        }

        public string ControlObject()
        {
            return "<p>Control Object Called.</p>";
        }

        private void button1_Click(object sender, EventArgs e)
        {
                // Call ActiveX
                myObject.SayHello("C# Button");
        }
    }
}
