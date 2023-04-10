namespace NorseTechnologies.AutoDocs.DocumentObjectModel
{
    public interface IRange
    {
        int Start { get; set; }
        int End { get; set; }
        AutoDocsStoryType StoryType { get; set; }
    }
}
