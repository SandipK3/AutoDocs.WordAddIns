using NorseTechnologies.AutoDocs.DocumentObjectModel;
using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.IO;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using System.Collections.Generic;
using System.Linq;

namespace NorseTechnologies.AutoDocs.MicrosoftWordDOM
{
    public class Bookmark : IBookmark
    {
        private Word.Document WordDoc { get; set; }
        public string Name { get; set; }
        public long Start { get; set; }
        public long End { get; set; }
        public string Text { get; set; }

        public void ClearBookmarkContents()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            WordDoc.Bookmarks[Name].Range.Text = String.Empty;
        }

        public void ContractRangeBeginning()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            int start = WordDoc.Bookmarks[Name].Range.Start;
            int end = WordDoc.Bookmarks[Name].Range.End;
            if (start < end)
            {
                start++;
                Range bookmarkRange = WordDoc.Bookmarks[Name].Range;
                if (WordDoc.Bookmarks[Name].Range.StoryType != WdStoryType.wdMainTextStory)
                {
                    bookmarkRange = WordDoc.Bookmarks[Name].Range;
                }
                bookmarkRange.Start = start;
                WordDoc.Bookmarks.Add(Name, bookmarkRange);
            }
        }

        public void ContractRangeEnd()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            int start = WordDoc.Bookmarks[Name].Range.Start;
            int end = WordDoc.Bookmarks[Name].Range.End;
            if (end > start)
            {
                end--;
                Range bookmarkRange = WordDoc.Bookmarks[Name].Range;
                if (WordDoc.Bookmarks[Name].Range.StoryType != WdStoryType.wdMainTextStory)
                {
                    bookmarkRange = WordDoc.Bookmarks[Name].Range;
                }
                bookmarkRange.End = end;
                WordDoc.Bookmarks.Add(Name, bookmarkRange);
            }
        }

        public void CopyBookmark(string destinationBookmark)
        {
            throw new NotImplementedException();
        }

        public void CopyEntityBookmarks(string bookmarkEntitySource, string bookmarkEntityDestination)
        {
            throw new NotImplementedException();
        }

        public void CopySpecificEntityBookmarks(IList<string> bookmarks, string bookmarkEntitySource, string bookmarkEntityDestination)
        {
            throw new NotImplementedException();
        }

        public void DeleteBookmarkContents()
        {
            throw new NotImplementedException();
        }

        public void DoubleBookmark(string bookmarkTarget)
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            // Warn is the target bookmark name already exists?

            WordDoc.Bookmarks.Add(bookmarkTarget, WordDoc.Bookmarks[Name].Range);
        }

        public void ExpandRangeBeginning()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            int start = WordDoc.Bookmarks[Name].Range.Start;
            int end = WordDoc.Bookmarks[Name].Range.End;
            if (start > 0)
            {
                start--;
                Range bookmarkRange = WordDoc.Bookmarks[Name].Range;
                if (WordDoc.Bookmarks[Name].Range.StoryType != WdStoryType.wdMainTextStory)
                {
                    bookmarkRange = WordDoc.Bookmarks[Name].Range;
                }
                bookmarkRange.Start = start;
                WordDoc.Bookmarks.Add(Name, bookmarkRange);
            }
        }

        public void ExpandRangeEnd()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            int start = WordDoc.Bookmarks[Name].Range.Start;
            int end = WordDoc.Bookmarks[Name].Range.End;
            if (end < WordDoc.Range().End)
            {
                end++;
                Range bookmarkRange = WordDoc.Bookmarks[Name].Range;
                if (WordDoc.Bookmarks[Name].Range.StoryType != WdStoryType.wdMainTextStory)
                {
                    bookmarkRange = WordDoc.Bookmarks[Name].Range;
                }
                bookmarkRange.End = end;
                WordDoc.Bookmarks.Add(Name, bookmarkRange);
            }
        }

        public string GetBookmarkContents()
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            return WordDoc.Bookmarks[Name].Range.Text;
        }

        public bool GetBookmarkContentsBoolean()
        {
            string[] booleanSetValues = { "true", "yes", "on" }; // This stinks for localization -- should pull these values from a string resource instead
            string bookmarkText = GetBookmarkContents().ToLower();
            bool value = booleanSetValues.Contains(bookmarkText);
            return value;
        }

        public int GetBookmarkContentsInteger()
        {
            int value = 0;
            if (Int32.TryParse(GetBookmarkContents(), out value))
                return value;
            return 0;
        }

        public IList<string> GetBookmarkNames(BookmarkSortOption sortOption)
        {
            throw new NotImplementedException();
        }

        public void RemoveInternalBookmarks()
        {
            throw new NotImplementedException();
        }

        public void SetBookmarkContents(string text)
        {
            if (null == WordDoc)
                throw new ArgumentNullException(nameof(WordDoc));

            if (!WordDoc.Bookmarks.Exists(Name))
                throw new ArgumentException("Bookmark with name [" + Name + " not found in document " + WordDoc.Name);

            WordDoc.Bookmarks[Name].Range.Text = text;
        }
    }
}
