using System;
using System.IO;

namespace DotMaysWind.Office
{
    public static class OfficeFileFactory
    {
        public static IOfficeFile CreateOfficeFile(String filePath)
        {
            String extension = Path.GetExtension(filePath).ToLower();

            if (String.Equals(".doc", extension))
            {
                return new WordBinaryFile(filePath);
            }
            else if (String.Equals(".ppt", extension))
            {
                return new PowerPointFile(filePath);
            }
            else if (String.Equals(".docx", extension))
            {
                return new WordOOXMLFile(filePath);
            }
            else if (String.Equals(".pptx", extension))
            {
                return new PowerPointOOXMLFile(filePath);
            }
            else
            {
                if (extension[extension.Length - 1] == 'x' || extension[extension.Length - 1] == 'X')
                {
                    return new OfficeOpenXMLFile(filePath);
                }
                else
                {
                    return new CompoundBinaryFile(filePath);
                }
            }
        }
    }
}