using System;
using System.Text;

namespace DotMaysWind.Office.Helper
{
    internal static class StringHelper
    {
        internal static String GetString(Boolean isUnicode, Byte[] data)
        {
            if (isUnicode)
            {
                return Encoding.Unicode.GetString(data);
            }
            else
            {
                return Encoding.GetEncoding("Windows-1252").GetString(data);
            }
        }

        internal static String ReplaceString(String origin)
        {
            StringBuilder sb = new StringBuilder();

            for (Int32 i = 0; i < origin.Length; i++)
            {
                if (origin[i] == '\r')
                {
                    sb.Append("\r\n");
                    continue;
                }

                if (origin[i] >= ' ')
                {
                    sb.Append(origin[i]);
                }
            }

            return sb.ToString();
        }
    }
}