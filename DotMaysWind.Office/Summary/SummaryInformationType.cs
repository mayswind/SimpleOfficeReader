using System;

namespace DotMaysWind.Office.Summary
{
    public enum SummaryInformationType : uint
    {
        Unknown = 0x00,
        CodePage = 0x01,
        Title = 0x02,
        Subject = 0x03,
        Author = 0x04,
        Keyword = 0x05,
        Commenct = 0x06,
        Template = 0x07,
        LastAuthor = 0x08,
        Reversion = 0x09,
        EditTime = 0x0A,
        CreateDateTime = 0x0C,
        LastSaveDateTime = 0x0D,
        PageCount = 0x0E,
        WordCount = 0x0F,
        CharCount = 0x10,
        ApplicationName = 0x12,
        Security = 0x13
    }
}