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

            Load += new EventHandler(Form1_Load);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.AllowWebBrowserDrop = false;
            webBrowser1.ObjectForScripting = this;
            webBrowser1.Url = new Uri(@"file:///c|\Users\jphillips\Documents\Visual Studio 2013\Projects\TestActiveX\TestActiveX\TestPage.html");

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
