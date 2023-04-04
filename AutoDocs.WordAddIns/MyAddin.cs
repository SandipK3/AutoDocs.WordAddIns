using Microsoft.Office.Core;
using NetOffice.Tools;
using NetOffice.WordApi;
using NetOffice.WordApi.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NetOffice.OfficeApi;
using Office = NetOffice.OfficeApi;
using System.Windows.Forms;
using ICTPFactory = NetOffice.OfficeApi.ICTPFactory;
using _CustomTaskPane = NetOffice.OfficeApi._CustomTaskPane;
using NetOffice.OfficeApi.Tools;
using Word = NetOffice.WordApi;
using Microsoft.Win32;
using NetOffice.VBIDEApi;
using System.Reflection;

namespace AutoDocs.WordAddIns
{
    [ComVisible(true)]
    [Guid("f7407235-8887-462b-94be-e916fd95b9b9")]
    [ProgId("AutoDocs.WordAddIns.MyAddin")]
    [COMAddin("MyAddin", "Addin description.", LoadBehavior.LoadAtStartup)]
    //[CustomPane(typeof(SampleControl), "AutoDocs WordAddIns", true, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 60, 60)]
    public class MyAddin : Word.Tools.COMAddin, IDisposable, ICustomTaskPaneConsumer
    {
        private static SampleControl _sampleControl;
        private bool _disposed = false;
        private static readonly string _prodId = "AutoDocs.TaskPaneAddin";

        private static Word.Application _wordApplication;
        internal static Word.Application Application { get { return _wordApplication; } }
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
        }

        private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _wordApplication.DocumentOpenEvent += Application_DocumentOpenEvent;
        }

        private void Application_DocumentOpenEvent(Document doc)
        {
            using (doc)
            {
                // start working with the document

            }
        }

        public void Dispose() => Dispose(true);
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Dispose managed state (managed objects).
                //_ctpFactory?.Dispose();
            }

            _disposed = true;
        }

        public void CTPFactoryAvailable(Microsoft.Office.Core.ICTPFactory CTPFactoryInst)
        {
            try
            {
                Office.ICTPFactory ctpFactory = new NetOffice.OfficeApi.ICTPFactory(_wordApplication, CTPFactoryInst);
                Office._CustomTaskPane taskPane = ctpFactory.CreateCTP(typeof(MyAddin).Assembly.GetName().Name + ".SampleControl", "AutoDocs TaskPane", Type.Missing);
                taskPane.DockPosition = (Office.Enums.MsoCTPDockPosition)MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = 300;
                taskPane.Visible = true;
                _sampleControl = taskPane.ContentControl as SampleControl;
                ctpFactory.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
