using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = NetOffice.WordApi;

namespace AutoDocs.WordAddIns
{
    [ComVisible(true)]
    [Guid("59CACEB4-C5F5-4BDF-9220-C7982E098B34")]
    [ProgId("AutoDocs.WordAddIns.AutoDocs365TaskPane")]
    public partial class AutoDocs365TaskPane : UserControl, Word.Tools.ITaskPane
    {
        Word.Application wordApp;

        Word.Document doc = null;
        bool visible = true;

        public AutoDocs365TaskPane()
        {
            InitializeComponent();
        }

        #region ITaskPane Implementation

        public void OnConnection(Word.Application application, _CustomTaskPane parentPane, object[] customArguments)
        {
            wordApp = application; 
        }

        public void OnDisconnection()
        {
            wordApp = null;
        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {
            // Do any layout or tasks that are dependent upon where the task pane is docked
        }

        public void OnVisibleStateChanged(bool visible)
        {
            this.visible = visible;
        }

        #endregion

        private void btnSaveDocument_Click(object sender, EventArgs e)
        {
            try
            {
                Word.Document doc = MyAddin.WordApplication.ActiveDocument;
                doc.SaveAs(@"C:\MyDocument.docx");
            }
            catch (Exception ex)
            {
            }
        }
    }
}
