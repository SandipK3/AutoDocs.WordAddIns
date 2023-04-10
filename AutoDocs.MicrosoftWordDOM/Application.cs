using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NorseTechnologies.AutoDocs.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.IO;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;

namespace NorseTechnologies.AutoDocs.MicrosoftWordDOM
{
    public class Application : IApplication
    {
        private Word.Application WordApp { get; set; }
        public IDocument ScrapDocument { get; set; }

        public IDocuments Documents { get; set; }

        public IDocument ActiveDocument 
        { 
            get
            {
                if ((null != WordApp) && (WordApp.Documents.Count > 0) && (null != WordApp.ActiveDocument))
                {
                    if (!Documents.Exists(WordApp.ActiveDocument.FullName))
                    {
                        (Documents as Documents).AddExisting(WordApp.ActiveDocument);
                    }
                    return Documents[WordApp.ActiveDocument.FullName];
                }
                return null;
            }
        }

        public Application()
        {

        }

        ~Application()
        {
            // When we destroy our Application object, we need to clean up any COM objects we created that we still have hanging around
            ReleaseScrapDocument();
            WordApp = null;
        }

        public void Initialize(object wordApplication)
        {
            WordApp = wordApplication as Word.Application;
            Documents = new Documents(this);
            (Documents as Documents).Initialize(WordApp);
        }

        public IDocument GetScrapDocument()
        {
            if (null == ScrapDocument)
            {
                ScrapDocument = Documents.Add(null, false, false);
                (ScrapDocument as Document).Initialize(WordApp, ScrapDocument);
            }
            return ScrapDocument;
        }

        public void ReleaseScrapDocument()
        {
            if (null != ScrapDocument)
            {
                ScrapDocument.Close(false);
                ScrapDocument = null;
            }
        }
    }
}
