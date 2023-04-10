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
    public class Document : IDocument
    {
        #region Private Properties
        private Word.Document WordDoc { get; set; }
        private Word.Application WordApp { get; set; }
        #endregion

        #region Public Properties

        public string Name 
        {
            get
            {
                if (null != WordDoc) { return WordDoc.Name; }
                return null;
            }
        }
        public string Path
        {
            get
            {
                if (null != WordDoc) { return WordDoc.Path; }
                return null;
            }
        }
        public string FullName { get { return System.IO.Path.Combine(Path, Name); } }
        public string AttachedTemplateName
        {
            get
            {
                if (null != WordDoc) 
                {
                    Word.Template template = WordDoc.AttachedTemplate as Template;
                    return template.Name;
                }
                return null;
            }
            set {  AttachedTemplateName = value; }
        }
        public string AttachedTemplatePath 
        {
            get
            {
                if (null != WordDoc) 
                {
                    Word.Template template = WordDoc.AttachedTemplate as Template;
                    return template.Path;
                }
                return null;
            }
            set { AttachedTemplatePath = value; }
        }
        public string AttachedTemplateFullName { get { return System.IO.Path.Combine(AttachedTemplatePath, AttachedTemplateName); } }

        public IApplication Application { get; set; }

        #endregion 

        public Document(IApplication wordApplication)
        {
            Application = wordApplication;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wordApp"></param>
        /// <param name="wordDocument"></param>
        /// 
        /// This is gross -- we have to have some way of setting up the parallel object references. For the Microsoft Word implementation of the DOM we need these, for a different implementation, we would need to do something different. I can safely make this cast here because I know that this IS the Word version of the AutoDocs DOM implementation.
        public void Initialize(object wordApp, object wordDocument)
        {
            WordApp = wordApp as Word.Application;
            WordDoc = wordDocument as Word.Document;
        }

        public void Activate()
        {
            if (null != WordDoc)
            {
                WordDoc.Activate();
            }
        }

        public void ClearBookmarkContentsGroup(string bookmarkBaseName)
        {
            throw new NotImplementedException();
        }

        public void Close(bool saveChanges)
        {
            if (null != WordDoc)
            {
                WordDoc.Close(saveChanges);
                WordDoc = null;
            }
            throw new NotImplementedException();
        }

        public IDocument Create(string templatePath, bool visible = true)
        {            
            throw new NotImplementedException();
        }

        public void GotoBookmark(string bookmarkName)
        {
            if (null != WordDoc)
            {
                if (WordDoc.Bookmarks.Exists(bookmarkName))
                {
                    WordDoc.Bookmarks[bookmarkName].Range.GoTo();
                }
            }
        }

        public IDocument Open(string filePath, bool visble = true, bool addToRecentFiles = true)
        {
            throw new NotImplementedException();
        }

        public IDocument Open(Stream stream, bool visble = true, bool addToRecentFiles = true)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            if (null != WordDoc)
            {
                WordDoc.Save();
            }
        }

        public void SaveAs(string filePath, AutoDocsDocumentFormat format)
        {
            if (null != WordDoc)
            {
                WordDoc.SaveAs2(filePath, WordFormatFromDocumentFormat(format));
            }
        }

        private WdSaveFormat WordFormatFromDocumentFormat(AutoDocsDocumentFormat autoDocsDocumentFormat) 
        {
            Dictionary<AutoDocsDocumentFormat, WdSaveFormat> autoDocsToWordFileFormatMap = new Dictionary<AutoDocsDocumentFormat, WdSaveFormat>()
            {
                { AutoDocsDocumentFormat.WordDocument, WdSaveFormat.wdFormatXMLDocument },
                { AutoDocsDocumentFormat.WordTemplate, WdSaveFormat.wdFormatXMLTemplate },
                { AutoDocsDocumentFormat.WordMacroEnabledTemplate, WdSaveFormat.wdFormatXMLDocumentMacroEnabled },
                { AutoDocsDocumentFormat.WordMacroEnabledDocument, WdSaveFormat.wdFormatXMLDocumentMacroEnabled },
                { AutoDocsDocumentFormat.Text, WdSaveFormat.wdFormatText }
            };

            if (autoDocsToWordFileFormatMap.ContainsKey(autoDocsDocumentFormat))
                return autoDocsToWordFileFormatMap[autoDocsDocumentFormat];

            return WdSaveFormat.wdFormatXMLDocument; // We don't know what format it is, let's just return .DOCX as a default
        }
    }
}
