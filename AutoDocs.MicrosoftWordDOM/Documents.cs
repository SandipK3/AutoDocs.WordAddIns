using NorseTechnologies.AutoDocs.DocumentObjectModel;
using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.IO;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using System.Collections.Generic;
using NetOffice.OfficeApi.Enums;
using System.Linq;

namespace NorseTechnologies.AutoDocs.MicrosoftWordDOM
{
    public class Documents : IDocuments
    {
        public IApplication Application { get; set; }
        private Word.Application WordApp { get; set; }

        public Dictionary<string, IDocument> documentCollection = new Dictionary<string, IDocument>();

        public IDocument this[int index]
        {
            get
            {
                if (index >= 0 && index < documentCollection.Count)
                    return documentCollection.ElementAt(index).Value;
                else
                    throw new IndexOutOfRangeException("The index '" + index + "' is out of range.");
            }
            set
            {
                if (index >= 0 && index < documentCollection.Count)
                {
                    string key = documentCollection.ElementAt(index).Key;
                    documentCollection[key] = value;
                }
                else
                    throw new IndexOutOfRangeException("The index '" + index + "' is out of range.");
            }
        }

        public IDocument this[string key]
        {
            get
            {
                if (documentCollection.ContainsKey(key))
                    return documentCollection[key];
                else
                    throw new KeyNotFoundException("The key " + key + " was not found in the Documents collection.");
            }
            set
            {
                if (!documentCollection.ContainsKey(key))
                    documentCollection[key] = (IDocument)value;
            }
        }

        public bool Exists(string filePath)
        {
            return documentCollection.ContainsKey(filePath);
        }

        public Documents(IApplication application)
        {
            Application = application;
        }

        /// <summary>
        /// Initialize is a method on the concrete class and NOT a part of the IDocuments interface. As such, we have to have knowledge of the fact that this particular version of the IDocuments is being utilized -- in this case, we have knowledge that the Microsoft Word version of the AutoDocs Document Object Model is in use due to the fact that this is used by the COM AddIn for Microsoft Word.
        /// </summary>
        /// <param name="application"></param>
        public void Initialize(Word.Application application)
        {
            WordApp = application;
        }

        public IDocument AddExisting(object document)
        {
            Word.Document wordDoc = document as Word.Document;
            IDocument docExisting = new AutoDocs.MicrosoftWordDOM.Document(Application);
            (docExisting as AutoDocs.MicrosoftWordDOM.Document).Initialize(WordApp, wordDoc);
            documentCollection[wordDoc.FullName] = docExisting;
            return docExisting;
        }

        public IDocument Add(string templatePath, bool template, bool visible)
        {
            Word.Document wordDoc = WordApp.Documents.Add(templatePath, template, WdNewDocumentType.wdNewBlankDocument, visible);
            IDocument docNew = new AutoDocs.MicrosoftWordDOM.Document(Application);
            (docNew as AutoDocs.MicrosoftWordDOM.Document).Initialize(WordApp, wordDoc); 
            documentCollection[wordDoc.FullName] = docNew;
            return docNew;
        }

        public IDocument Open(string filePath, bool readOnly = false, bool visible = true, bool addToRecentFiles = true, string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "", string writePasswordTemplate = "", bool openAndRepair = true, bool noEncodingDialog = false)
        {
            Word.Document wordDoc = WordApp.Documents.Open(filePath, false, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, WdOpenFormat.wdOpenFormatAuto, MsoEncoding.msoEncodingAutoDetect, visible, openAndRepair, WdDocumentDirection.wdLeftToRight, noEncodingDialog);
            IDocument docNew = new AutoDocs.MicrosoftWordDOM.Document(Application);
            (docNew as AutoDocs.MicrosoftWordDOM.Document).Initialize(WordApp, wordDoc); // This is gross -- we have to have some way of setting up the parallel object references. For the Microsoft Word implementation of the DOM we need these, for a different implementation, we would need to do something different. I can safely make this cast here because I know that this IS the Word version of the AutoDocs DOM implementation.
            documentCollection[wordDoc.FullName] = docNew;
            return docNew;
        }

        public IDocument Open(Stream stream, bool visble = true, bool addToRecentFiles = true)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            if (null != WordApp)
            {
                WordApp.Documents.Save();
            }
        }
    }
}
