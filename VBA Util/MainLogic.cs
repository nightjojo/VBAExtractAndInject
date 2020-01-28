using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Access = Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace VBA_Util
{
    abstract class MainLogic
    {
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowA(string lpClassName,string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowExA(IntPtr hWndParent, IntPtr hWndChildAfter,string lpszClass,string lpszWindow);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg,IntPtr wParam, string lParam);
        const uint WM_SETTEXT = 0x000C;
        const uint BM_CLICK = 0x00F5;

        protected enum TargetFileType
        {
            ACCESS,
            EXCEL
        }
        private static string _tgtfile;
        private static string _srcDir;
        private const char CR = '\r';
        private const char LF = '\n';
        private const char NULL = (char)0;
        protected int perc = 0;
        protected Access.Application AccApp = null;
        protected Excel.Application xlApp = null;
        protected void OpenApplication(string tgtFile, TargetFileType tgtFileType , string pwd="")
        {
            switch (tgtFileType)
            {
                case TargetFileType.ACCESS:
                    try
                    {
                        AccApp = new Access.Application();
                        //AccApp.Visible = false;
                        AccApp.OpenCurrentDatabase(tgtFile, true, pwd);
                        AccApp.Visible = false;
                        CommandBarControl ctrl = AccApp.VBE.CommandBars[1].FindControl(Id: 2578, Recursive: true);
                        Console.WriteLine(ctrl.Caption);
                        ctrl.Execute();
                        //find password dialog
                        string pjtname = null;
                        foreach (VBProject vbp in AccApp.VBE.VBProjects)
                        {
                            pjtname = vbp.Name;
                            break;
                        }
                        var hwnd = FindWindowA(null, pjtname + " Password");
                        if (hwnd == IntPtr.Zero) return;
                        //find password textbox
                        var hpwd = FindWindowExA(hwnd, IntPtr.Zero, "Edit", null);
                        // input password
                        SendMessage(hpwd, WM_SETTEXT, IntPtr.Zero, pwd);
                        // find OK button
                        var hbtn = FindWindowExA(hwnd, IntPtr.Zero, null, "OK");
                        if (hbtn != IntPtr.Zero)
                            SendMessage(hbtn, BM_CLICK, IntPtr.Zero, null);
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteExceptionLog(ex);
                    }
                    break;
                case TargetFileType.EXCEL:
                    // TODO
                    break;
            }
        }
        public abstract Boolean ProcessFile(string tgtFile, string srcDir, string pwd = "");
        protected void CloseApplication(TargetFileType tgtFileType)
        {
            switch (tgtFileType)
            {
                case TargetFileType.ACCESS:
                    // find orphaned property dialog
                    string pjtname = null;
                    foreach (VBProject vbp in AccApp.VBE.VBProjects)
                    {
                        pjtname = vbp.Name;
                        break;
                    }
                    var hwnd = FindWindowA(null, pjtname + " - Project Properties");
                    // find cancel button
                    var hbtn = FindWindowExA(hwnd, IntPtr.Zero, null, "Cancel");
                    if (hbtn != IntPtr.Zero)
                        SendMessage(hbtn, BM_CLICK, IntPtr.Zero, null);
                    if (AccApp != null)
                    {
                        AccApp.Quit(Access.AcQuitOption.acQuitSaveNone);
                    }
                    if (AccApp != null)
                    {
                        Marshal.ReleaseComObject(AccApp);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    break;
            }
        }
        public void SetFile(string tgtFile)
        {
            _tgtfile = tgtFile;
        }
        public void SetSourceDir(string srcDir)
        {
            _srcDir = srcDir;
        }
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