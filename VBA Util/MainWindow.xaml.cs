using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VBA_Util
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private static MainLogic logicMod = null;
        private BackgroundWorker _worker;
        private List<string> _errList = null;
        private string _target;
        private string _srcDir;
        private string _tab;
        private bool _isCancelled;
        private int percent;
        public int Percent
        {
            get { return this.percent; }
            set
            {
                this.percent = value;
                NotifyPropertyChange();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChange(string propertyName = "Percent")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public MainWindow()
        {
            InitializeComponent();
            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress = true;
            _worker.WorkerSupportsCancellation = true;
            _worker.DoWork += Work_DoWork;
            _worker.ProgressChanged += Work_ProgressChanged;
            _worker.RunWorkerCompleted += Work_Completed;
        }

        private ref MainLogic GetLogicInstance(string tabName)
        {
            logicMod = (MainLogic)(Activator.CreateInstance(null, "VBA_Util." + tabName).Unwrap());
            logicMod.SetFile(_target);
            logicMod.SetSourceDir(_srcDir);
            return ref logicMod;
        }

        //*****************************************************************
        //* Extract tab items
        //*****************************************************************
        private void ExtTgt_Drop(object sender, DragEventArgs e)
        {
            if (e.Data == null || e.Data.GetType() == null) return;
            var dataObj = e.Data.GetData(DataFormats.FileDrop);
            if (dataObj.GetType().IsArray && typeof(string).IsAssignableFrom(dataObj.GetType().GetElementType()))
            {
                var dataArr = (string[])dataObj;
                if (dataArr.Length > 1)
                {
                    MessageBox.Show("2つ以上のファイルを指定できません");
                    return;
                }
                if (!File.Exists(dataArr[0]))
                {
                    MessageBox.Show("無効なファイルパス");
                    return;
                }
                ExtTgtFile.Text = dataArr[0];
            }
        }

        private void ExtTgt_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void BtnExtTgt_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog { IsFolderPicker = false };
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                ExtTgtFile.Text = dialog.FileName;
            }
        }

        private void ExtOutDir_Drop(object sender, DragEventArgs e)
        {
            if (e.Data == null || e.Data.GetType() == null) return;
            var dataObj = e.Data.GetData(DataFormats.FileDrop);
            if (dataObj.GetType().IsArray && typeof(string).IsAssignableFrom(dataObj.GetType().GetElementType()))
            {
                var dataArr = (string[])dataObj;
                if (dataArr.Length > 1)
                {
                    MessageBox.Show("2つ以上のフォルダを指定できません");
                    return;
                }
                if (!Directory.Exists(dataArr[0]))
                {
                    MessageBox.Show("無効なフォルダパス");
                    return;
                }
                ExtOutDir.Text = dataArr[0];
            }
        }

        private void ExtOutDir_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void BtnOutDirExt_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog { IsFolderPicker = true };
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                ExtOutDir.Text = dialog.FileName;
            }
        }
        private void btnExec_Click(object sender, RoutedEventArgs e)
        {
            _isCancelled = false;
            btnExec.IsEnabled = false;
            //progressBar.Visibility = Visibility.Visible;
            //pgText.Visibility = Visibility.Visible;
            _tab = ((TabItem)tabCtrl.SelectedItem).Header.ToString();
            if (_tab.ToLower() == "extract")
            {
                _target = ExtTgtFile.Text;
                _srcDir = ExtOutDir.Text;
            }else
            {
                _target = InjTgtFile.Text;
                _srcDir = InjInDir.Text;
            }
            _worker.RunWorkerAsync();
        }
        //*****************************************************************
        //* Inject tab items
        //*****************************************************************
        private void InjTgtFile_Drop(object sender, DragEventArgs e)
        {
            if (e.Data == null || e.Data.GetType() == null) return;
            var dataObj = e.Data.GetData(DataFormats.FileDrop);
            if (dataObj.GetType().IsArray && typeof(string).IsAssignableFrom(dataObj.GetType().GetElementType()))
            {
                var dataArr = (string[])dataObj;
                if (dataArr.Length > 1)
                {
                    MessageBox.Show("2つ以上のファイルを指定できません");
                    return;
                }
                if (!File.Exists(dataArr[0]))
                {
                    MessageBox.Show("無効なファイルパス");
                    return;
                }
                InjTgtFile.Text = dataArr[0];
            }
        }

        private void InjTgtFile_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void BtnInjTgt_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog { IsFolderPicker = false };
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                InjTgtFile.Text = dialog.FileName;
            }
        }

        private void InjInDir_Drop(object sender, DragEventArgs e)
        {
            if (e.Data == null || e.Data.GetType() == null) return;
            var dataObj = e.Data.GetData(DataFormats.FileDrop);
            if (dataObj.GetType().IsArray && typeof(string).IsAssignableFrom(dataObj.GetType().GetElementType()))
            {
                var dataArr = (string[])dataObj;
                if (dataArr.Length > 1)
                {
                    MessageBox.Show("2つ以上のフォルダを指定できません");
                    return;
                }
                if (!Directory.Exists(dataArr[0]))
                {
                    MessageBox.Show("無効なフォルダパス");
                    return;
                }
                InjInDir.Text = dataArr[0];
            }
        }

        private void InjInDir_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void BtnInDirInj_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog { IsFolderPicker = true };
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                InjInDir.Text = dialog.FileName;
            }
        }

        //*****************************************************************
        //* Execute main logic with background worker
        //*****************************************************************
        private void Work_DoWork(object sender, DoWorkEventArgs e)
        {
            _worker.ReportProgress(0);
            String strErr = null;
            if (!File.Exists(_target)) strErr += "Input、";
            if (!Directory.Exists(_srcDir)) strErr += "Output";
            if (strErr != null)
            {
                if (strErr.Substring(strErr.Length - 1) == "、")
                    strErr = strErr.Substring(0, strErr.Length - 1);
                strErr += "パスが無効です";
                MessageBox.Show(strErr);
                _worker.CancelAsync();
                _isCancelled = true;
                return;
            }
            MainLogic mainLogic = GetLogicInstance(_tab);
            // acquire pwd from optional file(target filename plus ".pwd" without quotes)
            string pwd = "";
            if (File.Exists(_target + ".pwd"))
            {
                using (StreamReader sr = new StreamReader(_target + ".pwd", Encoding.UTF8))
                {
                    pwd = sr.ReadToEnd();
                }
            }
            if (_srcDir.Substring(_srcDir.Length - 2, 1) != "\\") _srcDir += "\\";
            if (!mainLogic.ProcessFile(_target, _srcDir, pwd))
            {
                MessageBox.Show("Error occured during process.\r\n See error log for detail");
                _worker.CancelAsync();
                _isCancelled = true;
                return;
            }
        }

        private void Work_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void Work_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!_isCancelled)
            {
                if (_errList == null)
                {
                    MessageBox.Show("Done");
                }
                else
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("Error processing the following file(s):");
                    foreach (var sFile in _errList)
                    {
                        sb.AppendLine(sFile);
                    }
                    MessageBox.Show(sb.ToString(), "Error");
                }
            }
            progressBar.Visibility = Visibility.Collapsed;
            pgText.Visibility = Visibility.Collapsed;
            btnExec.IsEnabled = true;
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            var binding = new Binding("Percent") { Source = this.Percent };
            progressBar.SetBinding(ProgressBar.ValueProperty, binding);
        }

    }
}
