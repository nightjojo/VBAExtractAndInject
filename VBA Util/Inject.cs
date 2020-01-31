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
    class Inject : MainLogic
    {
        public override bool ProcessFile(string tgtFile, string srcDir, string pwd="")
        {
            if (Regex.IsMatch(Path.GetExtension(tgtFile), ".*accd.*"))
            {
                Access.Application app = null;
                try
                {
                    OpenApplication(tgtFile, TargetFileType.ACCESS, pwd);
                    InjectCodeToAccess(ref app, srcDir);
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
        private bool InjectCodeToAccess(ref Access.Application app, string srcDir)
        {
            foreach (var srcFile in Directory.GetFiles(srcDir,"*",SearchOption.AllDirectories))
            {
                var fName = Path.GetFileNameWithoutExtension(srcFile);
                string srcContent = null;
                using (var sr = new StreamReader(srcFile, Encoding.UTF8))
                {
                    srcContent = sr.ReadToEnd();
                }
                bool isFound = false;
                foreach (VBProject vbp in app.VBE.VBProjects)
                {
                    foreach (VBComponent vbc in vbp.VBComponents)
                    {
                        CodeModule module = vbc.CodeModule;
                        if (module == null) continue;
                        //UPDATE module when found
                        if (module.Name == fName)
                        {
                            if (vbc.Type == vbext_ComponentType.vbext_ct_Document)
                            {
                                app.DoCmd.OpenForm(FormName: fName.Substring(5)
                                    , View: Access.AcFormView.acDesign
                                    , WindowMode: Access.AcWindowMode.acHidden);
                            }
                            else
                            {
                                app.DoCmd.OpenModule(module.Name);
                            }
                            //app.DoCmd.OpenModule(module.Name);
                            module.DeleteLines(1, module.CountOfLines);
                            module.AddFromString(srcContent);
                            if (vbc.Type == vbext_ComponentType.vbext_ct_Document)
                            {
                                app.DoCmd.Save(Access.AcObjectType.acForm, fName.Substring(5));
                            }
                            else
                            {
                                app.DoCmd.Save(Access.AcObjectType.acModule, module.Name);
                            }
                            //app.DoCmd.Close(Access.AcObjectType.acModule, module.Name, Access.AcCloseSave.acSaveYes);
                            isFound = true;
                            break;
                        }
                    }
                }
                
                // INSERT module when not found
                if (!isFound)
                {
                    //Get sub-directories
                    //Document,StdModule,ClassModule
                    var dir = Directory.GetParent(srcFile).Name;
                    var pjtName = Directory.GetParent(Directory.GetParent(srcFile).FullName).Name;
                    vbext_ComponentType moduleType;
                    if (Regex.IsMatch(dir, "StdModule"))
                        moduleType = vbext_ComponentType.vbext_ct_StdModule;
                    else if (Regex.IsMatch(dir, "ClassModule"))
                        moduleType = vbext_ComponentType.vbext_ct_ClassModule;
                    else continue;// DO NOT ADD form module via interop, use MS-ACCESS Export menu instead
                    foreach (VBProject pjt in app.VBE.VBProjects)
                    {
                        if (pjt.Name == pjtName)
                        {
                            VBComponent module = pjt.VBComponents.Add(moduleType);
                            module.Name = fName;
                            module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines);
                            module.CodeModule.AddFromString(srcContent);
                            app.DoCmd.Save(Access.AcObjectType.acModule, fName);
                            break;
                        }
                    }   
                }
            }
            return true;
        }
    }
}
