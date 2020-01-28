using System;
using System.IO;
using System.Text;
using Access = Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;

namespace VBA_Util
{
    class Extract : MainLogic
    {
        public override bool ProcessFile(string tgtFile, string srcDir, string pwd = "")
        {
            if (Regex.IsMatch(Path.GetExtension(tgtFile), ".*accd.*"))
            {
                AccApp = null;
                try
                {
                    OpenApplication(tgtFile, TargetFileType.ACCESS, pwd);
                    ExtractCodeFromAccess(ref AccApp, srcDir);
                }
                catch (Exception ex)
                {
                    Logger.WriteExceptionLog(ex);
                    return false;
                }
                finally
                {
                    CloseApplication(TargetFileType.ACCESS);
                }
            }
            else if (Regex.IsMatch(Path.GetExtension(tgtFile), ".*xls.*"))
            {
                //TODO implement later(?
            }
            else
            {
                // error invalid file extension
                return false;
            }
            return true;   
        }
        private bool ExtractCodeFromAccess(ref Access.Application AccApp, string srcDir)
        {
            foreach (VBProject vbp in AccApp.VBE.VBProjects)
            {
                foreach (VBComponent vbc in vbp.VBComponents)
                {
                    CreateTargetSourceFile(srcDir, vbp.Name,vbc);
                }
            }
            return true;
        }
        private void CreateTargetSourceFile(string srcDir, string pjtName, VBComponent vbc)
        {
            var module = vbc.CodeModule;
            if (module == null) return;
            string moduleType = null;
            switch (vbc.Type)
            {
                case vbext_ComponentType.vbext_ct_Document:
                    moduleType = "Document";
                    break;
                case vbext_ComponentType.vbext_ct_StdModule:
                    moduleType = "StdModule";
                    break;
                case vbext_ComponentType.vbext_ct_ClassModule:
                    moduleType = "ClassModule";
                    break;
            }
            string outputDir = srcDir + "\\" + pjtName + "\\" + moduleType;
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            using (var sw = new FileStream(outputDir + "\\" + module.Name + ".vb",
                            FileMode.Create, FileAccess.Write))
            {
                var contents = module.Lines[1, module.CountOfLines];
                sw.Write(Encoding.Unicode.GetBytes(contents), 0, Encoding.Unicode.GetByteCount(contents));
            }
        }
    }
}
