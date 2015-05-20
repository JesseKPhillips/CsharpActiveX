using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;


/// http://blogs.msdn.com/b/asiatech/archive/2011/12/05/how-to-develop-and-deploy-activex-control-in-c.aspx
/// http://stackoverflow.com/questions/11175145/create-com-activexobject-in-c-use-from-jscript-with-simple-event
///
/// Register with %NET64%\regasm /codebase <full path of dll file>
/// Unregister with %NET64%\regasm /u <full path of dll file>
namespace ActiveXObjectSpace
{

    /// <summary>
    /// Provides the ActiveX event listeners for Javascript.
    /// </summary>
    [Guid("4E250775-61A1-40B1-A57B-C7BBAA25F194"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IActiveXEvents
    {
        [DispId(1)]
        void OnUpdateString(string data);
    }

    /// <summary>
    /// Provides properties accessible from Javascript.
    /// </summary>
    [Guid("AAD0731A-E84A-48D7-B5F8-56FF1B7A61D3"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IActiveX
    {
        [DispId(10)]
        string CustomProperty { get; set; }
    }

    [ProgId("MyObject")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("7A5D58C7-1C27-4DFF-8C8F-F5876FF94C64")]
    [ComSourceInterfaces(typeof(IActiveXEvents))]
    public class MyObject : IActiveX
    {

        public delegate void OnContextChangeHandler(string data);
        public event OnContextChangeHandler OnUpdateString;

        // Dummy Method to use when firing the event
        private void MyActiveX_nMouseClick(string index)
        {

        }

        public MyObject()
        {
            // Bind event
            this.OnUpdateString = this.MyActiveX_nMouseClick;
        }

        [ComVisible(true)]
        public string CustomProperty { get; set; }


        [ComVisible(true)]
        public void SayHello(string who)
        {
            OnUpdateString("Calling Callback: " + who);
        }
    }
}