using System;

namespace DotMaysWind.Office.Helper
{
    internal static class BitHelper
    {
        internal static Boolean GetBitFromInteger(Int32 integer, Int32 bitIndex)
        {
            Int32 num = (Int32)Math.Pow(2, bitIndex);
            return (integer & num) == num;
        }
    }
}