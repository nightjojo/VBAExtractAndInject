using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace VBA_Util
{
    abstract class MainLogic
    {
        private static string _tgtfile;
        private static string _srcDir;
        private const char CR = '\r';
        private const char LF = '\n';
        private const char NULL = (char)0;
        protected int perc = 0;
        public void SetFile(string tgtFile)
        {
            _tgtfile = tgtFile;
        }
        public void SetSourceDir(string srcDir)
        {
            _srcDir = srcDir;
        }
        private static Encoding ms932 = Encoding.GetEncoding(932);
        private static Dictionary<string, string> dic = null;
        public abstract Boolean ProcessFile(string tgtFile, string srcDir);

        public static long CountLines(Stream stream)
        {
            var lineCount = 0L;
            var byteBuffer = new byte[1024 * 1024];
            var detectedEOL = NULL;
            var currentChar = NULL;

            int bytesRead;
            while ((bytesRead = stream.Read(byteBuffer, 0, byteBuffer.Length)) > 0)
            {
                for (var i = 0; i < bytesRead; i++)
                {
                    currentChar = (char)byteBuffer[i];

                    if (detectedEOL != NULL)
                    {
                        if (currentChar == detectedEOL)
                        {
                            lineCount++;
                        }
                    }
                    else if (currentChar == LF || currentChar == CR)
                    {
                        detectedEOL = currentChar;
                        lineCount++;
                    }
                }
            }

            // We had a NON-EOL character(EOF) at the end without a new line
            if (currentChar != LF && currentChar != CR && currentChar != NULL)
            {
                lineCount++;
            }
            return lineCount;
        }
    }
}