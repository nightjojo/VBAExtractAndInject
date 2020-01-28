using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBA_Util
{
    class Logger
    {
        public static void WriteExceptionLog(Exception ex, string logPath = null)
        {
            if (logPath == null || !Directory.Exists(Path.GetDirectoryName(logPath)))
            {
                logPath = Directory.GetCurrentDirectory() + @"\errors.log";
            }
            using (var sw = new FileStream(logPath,FileMode.Append, FileAccess.Write))
            {
                var sb = new StringBuilder();
                sb.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":");
                sb.AppendLine("HResult: " + ex.HResult);
                sb.AppendLine(ex.Message);
                sb.AppendLine(ex.StackTrace);
                sw.Write(Encoding.Unicode.GetBytes(sb.ToString()), 0, Encoding.Unicode.GetByteCount(sb.ToString()));
            }
        }
    }
}
