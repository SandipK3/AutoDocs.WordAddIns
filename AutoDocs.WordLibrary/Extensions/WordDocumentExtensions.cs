using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System;

namespace NorseTechnologies.AutoDocs.WordLibrary.Extensions
{
    public static class WordDocumentExtensions
    {
        public static void ForAllRanges(this Document doc, Action<Range> execute)
        {
            if (null == doc)
                throw new ArgumentNullException(nameof(doc));

            if (null == execute)
                throw new ArgumentNullException(nameof(execute));

            long longLink;
            Range range;

            longLink = (long)doc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
            foreach (Range rangeStory in doc.StoryRanges)
            {
                range = rangeStory;

                execute(range);

                // Iterate through all linked stories
                range = range.NextStoryRange;
                while (range != null)
                {
                    execute(range);

                    switch (range.StoryType)
                    {
                        case WdStoryType.wdEvenPagesHeaderStory:
                        case WdStoryType.wdPrimaryHeaderStory:
                        case WdStoryType.wdEvenPagesFooterStory:
                        case WdStoryType.wdPrimaryFooterStory:
                        case WdStoryType.wdFirstPageHeaderStory:
                        case WdStoryType.wdFirstPageFooterStory:

                            if (range.ShapeRange.Count > 0)
                            {
                                foreach (Shape shape in range.ShapeRange)
                                {
                                    if (shape.TextFrame.HasText == -1)
                                    {
                                        execute(shape.TextFrame.TextRange);
                                    }
                                }
                            }
                            break;
                    }
                    range = range.NextStoryRange;
                }
            }
        }
    }
}
