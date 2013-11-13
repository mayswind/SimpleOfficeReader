using System;

namespace DotMaysWind.Office
{
    public interface IWordFile
    {
        String ParagraphText { get; }

        String HeaderAndFooterText { get; }

        String CommentText { get; }

        String FootnoteText { get; }

        String EndnoteText { get; }
    }
}