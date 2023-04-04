using NetOffice.WordApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = NetOffice.WordApi;

namespace AutoDocs.WordAddIns
{
    public partial class SampleControl : UserControl
    {
        Word.Application wordApp = new Word.Application();
        Word.Document doc = null;
        public SampleControl()
        {
            InitializeComponent();
        }

        private void btnSaveDocument_Click(object sender, EventArgs e)
        {
            try
            {
                Word.Document doc = MyAddin.Application.ActiveDocument;
                doc.SaveAs(@"C:\MyDocument.docx");
                doc.Dispose();
                wordApp.Dispose();
            }
            catch (Exception ex)
            {
                doc.Dispose();
                wordApp.Dispose();
            }
        }
    }
}
