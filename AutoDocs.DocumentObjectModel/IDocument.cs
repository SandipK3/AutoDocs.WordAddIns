using System.IO;

namespace NorseTechnologies.AutoDocs.DocumentObjectModel
{
    public interface IDocument
    {
        ///////////////////////////////////////////////////////////////////////
        // IWordApplication
        ///////////////////////////////////////////////////////////////////////
        IApplication Application { get; set; }

        ///////////////////////////////////////////////////////////////////////
        // Document Properties
        ///////////////////////////////////////////////////////////////////////
        string Name { get; }
        string Path { get; }
        string FullName { get; }
        string AttachedTemplateName { get; set; }
        string AttachedTemplatePath { get; set; }
        string AttachedTemplateFullName { get; }

        ///////////////////////////////////////////////////////////////////////
        // Document Methods
        ///////////////////////////////////////////////////////////////////////
        void Activate();
        void Save();
        void SaveAs(string filePath, AutoDocsDocumentFormat format);
        void Close(bool saveChanges);

        ///////////////////////////////////////////////////////////////////////
        // Bookmark Methods
        ///////////////////////////////////////////////////////////////////////
        void GotoBookmark(string bookmarkName);
        void ClearBookmarkContentsGroup(string bookmarkBaseName);
    }
}