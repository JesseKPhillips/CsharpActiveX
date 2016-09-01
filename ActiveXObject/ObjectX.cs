using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;


/// http://blogs.msdn.com/b/asiatech/archive/2011/12/05/how-to-develop-and-deploy-activex-control-in-c.aspx
/// http://stackoverflow.com/questions/11175145/create-com-activexobject-in-c-use-from-jscript-with-simple-event
/// http://stackoverflow.com/questions/421857/using-activex-propertybags-from-c-sharp
///
/// Register with & "$env:windir\Microsoft.NET\Framework\v4.0.30319\regasm" /codebase <full path of dll file>
/// Unregister with & "$env:windir\Microsoft.NET\Framework\v4.0.30319\regasm" /u <full path of dll file>
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
    
    [ComVisible(true), ComImport,
    Guid("37D84F60-42CB-11CE-8135-00AA004BB851"),//Guid("5738E040-B67F-11d0-BD4D-00A0C911CE86"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistPropertyBag //: IPersist
    {
        #region IPersist
        [PreserveSig]
        /*new*/ int GetClassID([Out] out Guid pClassID);
        #endregion

        [PreserveSig]
        int InitNew();

        [PreserveSig]
        int Load(
        [In] IPropertyBag pPropBag,
        [In, MarshalAs(UnmanagedType.Interface)] object pErrorLog
        );

        [PreserveSig]
        int Save(
        IPropertyBag pPropBag,
        [In, MarshalAs(UnmanagedType.Bool)] bool fClearDirty,
        [In, MarshalAs(UnmanagedType.Bool)] bool fSaveAllProperties
        );
    }

    [ComVisible(true), ComImport,
    Guid("55272A00-42CB-11CE-8135-00AA004BB851"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyBag
    {
        [PreserveSig]
        int Read(
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
        [In, Out, MarshalAs(UnmanagedType.Struct)]    ref    object pVar,
        [In] IntPtr pErrorLog);

        [PreserveSig]
        int Write(
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
        [In, MarshalAs(UnmanagedType.Struct)] ref object pVar);
    }



    [ProgId("MyObject")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("7A5D58C7-1C27-4DFF-8C8F-F5876FF94C64")]
    [ComSourceInterfaces(typeof(IActiveXEvents))]
    public class MyObject : IActiveX, IPersistPropertyBag, IPropertyBag
    {

        public delegate void OnContextChangeHandler(string data);
        public event OnContextChangeHandler OnUpdateString;

        [ComVisible(true)]
        public string CustomProperty { get; set; }


        [ComVisible(true)]
        public void SayHello(string who)
        {
            OnUpdateString(who + "  :CustomProperty=" + this.CustomProperty);
            MessageBox.Show(" public void SayHello(string who):" + who + "  :CustomProperty=" + this.CustomProperty);
        }



        #region IPropertyBag Members

        public int Read(string pszPropName, ref object pVar, IntPtr pErrorLog)
        {
            pVar = null;
            switch (pszPropName)
            {
                case "CustomProperty": pVar = this.CustomProperty; break;
               
            }

            return 0;
        }

        public int Write(string pszPropName, ref object pVar)
        {
            switch (pszPropName)
            {
                case "CustomProperty": this.CustomProperty = (string)pVar; break;
            }

            return 0;
        }

        #endregion

        #region IPersistPropertyBag Members

        public int GetClassID(out Guid pClassID)
        {
            throw new NotImplementedException();
        }

        public int InitNew()
        {
            return 0;
        }

        public int Load(IPropertyBag pPropBag, object pErrorLog)
        {
            object val = null;

            pPropBag.Read("CustomProperty", ref val, IntPtr.Zero);
            Write("CustomProperty", ref val);
            return 0;
        }

        public int Save(IPropertyBag pPropBag, bool fClearDirty, bool fSaveAllProperties)
        {
            return 0;
        }

        #endregion

    }
}