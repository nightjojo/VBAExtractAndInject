using System;
using System.IO;
using System.Text;
using Access = Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;

namespace VBA_Util
{
    class Extract : MainLogic
    {
        public override bool ProcessFile(string tgtFile, string srcDir)
        {
            if (Regex.IsMatch(Path.GetExtension(tgtFile), ".*accd.*"))
            {
                Access.Application app=null;
                Dao.Database db=null;
                try
                {
                    app = new Access.Application();
                    app.OpenCurrentDatabase(tgtFile);
                    db = app.CurrentDb();
                    ExtractCodeFromAccess(ref app, srcDir);
                }
                catch (Exception ex)
                {
                    using (var sw = new FileStream(Directory.GetCurrentDirectory() + @"\errors.log",
                                        FileMode.Append, FileAccess.Write))
                    {
                        var sb = new StringBuilder();
                        sb.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":");
                        sb.AppendLine("HResult: " + ex.HResult);
                        sb.AppendLine(ex.Message);
                        sb.AppendLine(ex.StackTrace);
                        sw.Write(Encoding.Unicode.GetBytes(sb.ToString()), 0, Encoding.Unicode.GetByteCount(sb.ToString()));
                    }
                    return false;
                }
                finally
                {
                    if (db != null)
                    {
                        db.Close();
                    }
                    if (app != null)
                    {
                        app.Quit();
                    }
                    if (db != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(db);
                        if (app != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
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
        private bool ExtractCodeFromAccess(ref Access.Application app, string srcDir)
        {
            foreach (VBProject vbp in app.VBE.VBProjects)
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
