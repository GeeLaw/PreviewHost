using System;
using System.Runtime.InteropServices;

namespace PreviewHost.Interop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct LogFontW
    {

        public int FontHeight;
        public int FontWidth;
        public int FontEscapement;
        public int FontOrientation;
        public int Weight;
        public byte Italic;
        public byte Underline;
        public byte StrikeOut;
        public byte CharSet;
        public byte OutPrecision;
        public byte ClipPrecision;
        public byte Quality;
        public byte PitchAndFamily;

        char FaceName0;
        char FaceName1;
        char FaceName2;
        char FaceName3;
        char FaceName4;
        char FaceName5;
        char FaceName6;
        char FaceName7;
        char FaceName8;
        char FaceName9;
        char FaceName10;
        char FaceName11;
        char FaceName12;
        char FaceName13;
        char FaceName14;
        char FaceName15;
        char FaceName16;
        char FaceName17;
        char FaceName18;
        char FaceName19;
        char FaceName20;
        char FaceName21;
        char FaceName22;
        char FaceName23;
        char FaceName24;
        char FaceName25;
        char FaceName26;
        char FaceName27;
        char FaceName28;
        char FaceName29;
        char FaceName30;
        char FaceName31;

        public string FaceName
        {
            get
            {
                char[] charArray = new char[] { FaceName0, FaceName1, FaceName2, FaceName3, FaceName4, FaceName5, FaceName6, FaceName7, FaceName8, FaceName9, FaceName10, FaceName11, FaceName12, FaceName13, FaceName14, FaceName15, FaceName16, FaceName17, FaceName18, FaceName19, FaceName20, FaceName21, FaceName22, FaceName23, FaceName24, FaceName25, FaceName26, FaceName27, FaceName28, FaceName29, FaceName30, FaceName31 };
                int length;
                for (length = 0; length < charArray.Length; ++length)
                    if (charArray[length] == '\0')
                        break;
                if (length == charArray.Length)
                    return null;
                return new string(charArray, 0, length);
            }
            set
            {
                int length = value.Length;
                if (length > 31)
                    throw new ArgumentOutOfRangeException("value", "The string is too long.");
                FaceName0 = length >= 0 ? value[0] : '\0';
                FaceName1 = length >= 1 ? value[1] : '\0';
                FaceName2 = length >= 2 ? value[2] : '\0';
                FaceName3 = length >= 3 ? value[3] : '\0';
                FaceName4 = length >= 4 ? value[4] : '\0';
                FaceName5 = length >= 5 ? value[5] : '\0';
                FaceName6 = length >= 6 ? value[6] : '\0';
                FaceName7 = length >= 7 ? value[7] : '\0';
                FaceName8 = length >= 8 ? value[8] : '\0';
                FaceName9 = length >= 9 ? value[9] : '\0';
                FaceName10 = length >= 10 ? value[10] : '\0';
                FaceName11 = length >= 11 ? value[11] : '\0';
                FaceName12 = length >= 12 ? value[12] : '\0';
                FaceName13 = length >= 13 ? value[13] : '\0';
                FaceName14 = length >= 14 ? value[14] : '\0';
                FaceName15 = length >= 15 ? value[15] : '\0';
                FaceName16 = length >= 16 ? value[16] : '\0';
                FaceName17 = length >= 17 ? value[17] : '\0';
                FaceName18 = length >= 18 ? value[18] : '\0';
                FaceName19 = length >= 19 ? value[19] : '\0';
                FaceName20 = length >= 20 ? value[20] : '\0';
                FaceName21 = length >= 21 ? value[21] : '\0';
                FaceName22 = length >= 22 ? value[22] : '\0';
                FaceName23 = length >= 23 ? value[23] : '\0';
                FaceName24 = length >= 24 ? value[24] : '\0';
                FaceName25 = length >= 25 ? value[25] : '\0';
                FaceName26 = length >= 26 ? value[26] : '\0';
                FaceName27 = length >= 27 ? value[27] : '\0';
                FaceName28 = length >= 28 ? value[28] : '\0';
                FaceName29 = length >= 29 ? value[29] : '\0';
                FaceName30 = length >= 30 ? value[30] : '\0';
                FaceName31 = '\0';
            }
        }
    }
}
