using System.IO;

namespace NorseTechnologies.AutoDocs.DocumentObjectModel
{
    public interface IDocuments 
    {        
        IDocument this[int index] { get; set; }
        IDocument this[string key] { get; set; }
        
        /// <summary>
        /// Save will perform a save operation on all documents open in the collection of documents.
        /// </summary>
        IDocument Add(string templatePath, bool template, bool visible);
        bool Exists(string filePath);
        IDocument Open(string filePath, bool readOnly = false, bool visible = true, bool addToRecentFiles = true, string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "", string writePasswordTemplate = "", bool openAndRepair = true, bool noEncodingDialog = false);
        IDocument Open(Stream stream, bool visble = true, bool addToRecentFiles = true);
        void Save();
    }
}
