using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using NetOffice.Tools;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
using NorseTechnologies.AutoDocs.DocumentObjectModel;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;

namespace AutoDocs.WordAddIns
{
    [ComVisible(true)]
    [Guid("f7407235-8887-462b-94be-e916fd95b9b9")]
    [ProgId("AutoDocs.WordAddIns.MyAddin")]
    [COMAddin("MyAddin", "Addin description.", LoadBehavior.LoadAtStartup)]
    [CustomUI("RibbonUI.xml", true)]
    public class MyAddin : Word.Tools.COMAddin, IDisposable
    {
        private static AutoDocs365TaskPane _sampleControl;
        private bool _disposed = false;
        private static readonly string _prodId = "AutoDocs.TaskPaneAddin";
        ICTPFactory _ctpFactory = null;
        _CustomTaskPane taskPane = null;
        IApplication AutoDocsApplication { get; set; }

        private static Word.Application _wordApplication;
        internal static Word.Application WordApplication { get { return _wordApplication; } }
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
            this.OnDisconnection += MyAddin_OnDisconnection;
        }

        #region Event Handlers

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
            AutoDocsApplication = new NorseTechnologies.AutoDocs.MicrosoftWordDOM.Application();
            AutoDocsApplication.Initialize(_wordApplication);
        }

        private void WordApplication_DocumentSyncEvent(Document doc, MsoSyncEventType syncEventType)
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
            IDocument activeDocument = AutoDocsApplication.ActiveDocument;
        }

        private void WordApplication_DocumentOpenEvent(Document doc)
        {
            using (doc)
            {
                // start working with the document

            }
        }

        #endregion

        #region IDisposable

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

        #endregion 


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
                tpi.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                tpi.Width = 460;
                tpi.Visible = true;

                taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                taskPane.Width = 460;
                taskPane.Visible = true;
            }
            catch (Exception ex)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Ribbon Customization

        public Bitmap RibbonLoadImage(string imageName)
        {
            switch (imageName)
            {
                case "section16x16.png":
                    return new Bitmap(Properties.Resources.SectionSymbol16x16);
            }
            return null;
        
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            if (null == control)
                return null;

            switch (control.Id)
            {
                case "section16x16.png":
                    return new Bitmap(Properties.Resources.SectionSymbol16x16);
                default:
                    return null;
            }
        }

        public bool IsContentLibraryVisible(IRibbonControl control)
        {
            return true;
        }

        public bool IsTemplateLibraryVisible(IRibbonControl control)
        {
            return true;
        }

        public bool IsAutoDocsEnterpriseVisible(IRibbonControl control)
        {
            return true;
        }
                
        public bool IsAutoDocs365Admin(IRibbonControl control)
        {
            return true;
        }
        #endregion

        public void InsertSectionSymbolClick(IRibbonControl control)
        {
            if (null == control)
                return;

            InsertCharacter("Section Symbol", "\u00A7");
        }

        public void InsertCharacter(string characterName, string character)
        {
            if (null == WordApplication.ActiveDocument)
                return;

            if ((WordApplication.Selection.End - WordApplication.Selection.Start) > 1)
            {
                MessageBox.Show(string.Format("You must have an insertion point selection in the document template to insert the {0} character.", characterName));
                return;
            }

            if (WordApplication.Selection.End == WordApplication.Selection.Start)
            {
                WordApplication.Selection.InsertAfter(character);
            }
            else
            {
                WordApplication.Selection.Text = character;
            }

            WordApplication.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        }

        // If we create a Plugin Interface for the various AutoDocs 365 tools, we might use a single onAction method and then pass the IRibbonControl instance to the plugin manager to dispatch the command to all the listening plugins to decide who needs to process the message. This would allow us to have containment for each application in a separate component instead of building one large monolithic app.
        public void AutoDocs365RibbonButtonClick(IRibbonControl control)
        {
            // Route this click message to the Plugin Manager to notify listening plugins that a button was clicked and give them an opportunity to handle the message
        }

        public void SearchContentLibraryButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void SubmitContentLibraryButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void ContentManagementContentLibraryButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void ContentLibrarySettingsButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void SearchTemplateLibraryButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void TemplateLibraryInsertDataField(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void TemplateLibraryConditionalContent(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void ContentManagementTemplateLibraryButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void TemplateLibrarySettingsButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void AutoDocsEnterpriseSettingsButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void AutoDocs365SettingsButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }

        public void AboutAutoDocs365ButtonClick(IRibbonControl control)
        {
            MessageBox.Show(string.Format("{0} button clicked.", control.Id));
        }
    }
}
