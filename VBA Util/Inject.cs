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
                try
                {
                    OpenApplication(tgtFile, TargetFileType.ACCESS, pwd);
                    InjectCodeToAccess(srcDir);
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
        private bool InjectCodeToAccess(string srcDir)
        {
            // get project name(s)
            foreach (VBProject vbp in AccApp.VBE.VBProjects)
            {
                var pjtName = vbp.Name;
                // get project directory
                var pjtDir = Directory.GetDirectories(srcDir, pjtName);
                // skip when not found
                if (pjtDir == null || pjtDir.Length == 0) continue;

                foreach (var srcFile in Directory.GetFiles(pjtDir[0], "*", SearchOption.AllDirectories))
                {
                    var fName = Path.GetFileNameWithoutExtension(srcFile);
                    if (fName == null || fName == "") continue;
                    string srcContent = null;
                    using (var sr = new StreamReader(srcFile, Encoding.UTF8))
                    {
                        srcContent = sr.ReadToEnd();
                    }
                    bool isFound = false;
                    foreach (VBComponent vbc in vbp.VBComponents)
                    {
                        CodeModule module = vbc.CodeModule;
                        if (module == null) continue;
                        //UPDATE module when found
                        if (module.Name == fName)
                        {
                            if (vbc.Type == vbext_ComponentType.vbext_ct_Document)
                            {
                                AccApp.DoCmd.OpenForm(FormName: fName.Substring(5)
                                    , View: Access.AcFormView.acDesign
                                    , WindowMode: Access.AcWindowMode.acHidden);
                                //AccApp.DoCmd.OpenForm(FormName: module.Name
                                //    , View: Access.AcFormView.acDesign
                                //    , WindowMode: Access.AcWindowMode.acHidden);
                            }
                            else
                            {
                                AccApp.DoCmd.OpenModule(module.Name);
                            }
                            //AccApp.DoCmd.OpenModule(module.Name);
                            module.DeleteLines(1, module.CountOfLines);
                            module.AddFromString(srcContent);
                            if (vbc.Type == vbext_ComponentType.vbext_ct_Document)
                            {
                                AccApp.DoCmd.Save(Access.AcObjectType.acForm, fName.Substring(5));
                                //AccApp.DoCmd.Save(Access.AcObjectType.acForm, module.Name);
                            }
                            else
                            {
                                AccApp.DoCmd.Save(Access.AcObjectType.acModule, module.Name);
                            }
                            //AccApp.DoCmd.Close(Access.AcObjectType.acModule, module.Name, Access.AcCloseSave.acSaveYes);
                            isFound = true;
                            break;
                        }
                    }
                    
                    // INSERT module when not found
                    if (!isFound)
                    {
                        //Get sub-directories
                        //Document,StdModule,ClassModule
                        var dir = Directory.GetParent(srcFile).Name;
                        //var pjtName = Directory.GetParent(Directory.GetParent(srcFile).FullName).Name;
                        vbext_ComponentType moduleType;
                        if (Regex.IsMatch(dir, "StdModule"))
                            moduleType = vbext_ComponentType.vbext_ct_StdModule;
                        else if (Regex.IsMatch(dir, "ClassModule"))
                            moduleType = vbext_ComponentType.vbext_ct_ClassModule;
                        else continue;// DO NOT ADD form module via interop, use MS-ACCESS Export menu instead
                        foreach (VBProject pjt in AccApp.VBE.VBProjects)
                        {
                            if (pjt.Name == pjtName)
                            {
                                VBComponent module = pjt.VBComponents.Add(moduleType);
                                module.Name = fName;
                                module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines);
                                module.CodeModule.AddFromString(srcContent);
                                AccApp.DoCmd.Save(Access.AcObjectType.acModule, fName);
                                break;
                            }
                        }
                    }
                }

            }
            return true;
        }
    }
}
