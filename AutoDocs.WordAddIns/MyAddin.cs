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
    public class MyAddin : Word.Tools.COMAddin, IDisposable
    {
        private static AutoDocs365TaskPane _sampleControl;
        private bool _disposed = false;
        private static readonly string _prodId = "AutoDocs.TaskPaneAddin";
        ICTPFactory _ctpFactory = null;
        _CustomTaskPane taskPane = null;

        private static Word.Application _wordApplication;
        internal static Word.Application WordApplication { get { return _wordApplication; } }
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
            this.OnDisconnection += MyAddin_OnDisconnection;
        }

        private void MyAddin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _ctpFactory.Dispose();
            _ctpFactory = null;

            if (null != taskPane)
            {
                taskPane.Dispose();
                taskPane = null;
            }
        }

        private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _wordApplication = application as Word.Application;
            WordApplication.DocumentOpenEvent += WordApplication_DocumentOpenEvent;
            WordApplication.DocumentChangeEvent += WordApplication_DocumentChangeEvent;
            WordApplication.DocumentBeforeCloseEvent += WordApplication_DocumentBeforeCloseEvent;
            WordApplication.DocumentBeforePrintEvent += WordApplication_DocumentBeforePrintEvent;
            WordApplication.DocumentSyncEvent += WordApplication_DocumentSyncEvent;
        }

        private void WordApplication_DocumentSyncEvent(Document doc, Office.Enums.MsoSyncEventType syncEventType)
        {
        }

        private void WordApplication_DocumentBeforePrintEvent(Document doc, ref bool cancel)
        {
        }

        private void WordApplication_DocumentBeforeCloseEvent(Document doc, ref bool cancel)
        {
        }

        private void WordApplication_DocumentChangeEvent()
        {
            CreateAutoDocs365TaskPane(); // Create an AutoDocs 365 Custom Task Pane if one doesn't already exist
        }

        private void WordApplication_DocumentOpenEvent(Document doc)
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

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            _ctpFactory = new ICTPFactory(this.Application, CTPFactoryInst);
        }

        private void CreateAutoDocs365TaskPane()
        {
            if (null == _ctpFactory)
                return;

            if (null != taskPane)
                return;

            try
            {
                taskPane = _ctpFactory.CreateCTP("AutoDocs.WordAddIns.AutoDocs365TaskPane", "AutoDocs 365");
                TaskPaneInfo tpi = TaskPanes.Add(typeof(AutoDocs365TaskPane), "AutoDocs 365");
                tpi.DockPosition = (Office.Enums.MsoCTPDockPosition)MsoCTPDockPosition.msoCTPDockPositionLeft;
                tpi.Width = 460;
                tpi.Visible = true;

                taskPane.DockPosition = (Office.Enums.MsoCTPDockPosition)MsoCTPDockPosition.msoCTPDockPositionLeft;
                taskPane.Width = 460;
                taskPane.Visible = true;
            }
            catch (Exception ex)
            {
            }
        }
    }
}
