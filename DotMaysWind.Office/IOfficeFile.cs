using System;
using System.Collections.Generic;

using DotMaysWind.Office.Summary;

namespace DotMaysWind.Office
{
    public interface IOfficeFile
    {
        Dictionary<String, String> DocumentSummaryInformation { get; }

        Dictionary<String, String> SummaryInformation { get; }
    }
}