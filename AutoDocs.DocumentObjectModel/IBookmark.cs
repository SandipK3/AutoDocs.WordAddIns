using System.Collections.Generic;

namespace NorseTechnologies.AutoDocs.DocumentObjectModel
{
    public interface IBookmark
    {
        ///////////////////////////////////////////////////////////////////////
        // Bookmark Properties
        ///////////////////////////////////////////////////////////////////////
        string Name { get; set; }
        long Start { get; set; }
        long End { get; set; }
        string Text { get; set; }
        ///////////////////////////////////////////////////////////////////////
        // Bookmark Methods
        ///////////////////////////////////////////////////////////////////////
        int GetBookmarkContentsInteger();
        string GetBookmarkContents();
        bool GetBookmarkContentsBoolean();
        void SetBookmarkContents(string text);
        void ClearBookmarkContents();
        void DeleteBookmarkContents();
        void DoubleBookmark(string bookmarkTarget);
        void CopyBookmark(string destinationBookmark);
        void RemoveInternalBookmarks();
        void CopyEntityBookmarks(string bookmarkEntitySource, string bookmarkEntityDestination);
        void CopySpecificEntityBookmarks(IList<string> bookmarks, string bookmarkEntitySource, string bookmarkEntityDestination);
        IList<string> GetBookmarkNames(BookmarkSortOption sortOption);

        // Bookmark Scootching Methods
        void ExpandRangeBeginning();

        void ContractRangeBeginning();

        void ExpandRangeEnd();

        void ContractRangeEnd();
    }
}
