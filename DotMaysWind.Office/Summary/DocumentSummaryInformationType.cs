using System;

namespace DotMaysWind.Office.Summary
{
    public enum DocumentSummaryInformationType : uint
    {
        Unknown = 0x00,
        CodePage = 0x01,
        Category = 0x02,
        PresentationTarget = 0x03,
        Bytes = 0x04,
        LineCount = 0x05,
        ParagraphCount = 0x06,
        Slides = 0x07,
        Notes = 0x08,
        HiddenSlides = 0x09,
        MMClips = 0x0A,
        Scale = 0x0B,
        HeadingPairs = 0x0C,
        DocumentParts = 0x0D,
        Manager = 0x0E,
        Company = 0x0F,
        LinksDirty = 0x10,
        CountCharsWithSpaces = 0x11,
        SharedDoc = 0x13,
        HyperLinksChanged = 0x16,
        Version = 0x17,
        ContentStatus = 0x1B
    }
}