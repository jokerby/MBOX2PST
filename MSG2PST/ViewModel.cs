using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using MSG2PST.Annotations;
using Redemption;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace MSG2PST
{
    public sealed class ViewModel : INotifyPropertyChanged
    {
        private readonly RDOSession _session = new RDOSession();
        private string _destinationPath;
        private string _sourcePath;
        private Thread _thread;
        private DateTime _start;
        private string _message = "Finished";
        private string _currentFile = string.Empty;
        private bool _isRun;
        private readonly List<Task> _tasks = new List<Task>();

        public string CurrentFile
        {
            get { return _currentFile; }
            set
            {
                if (value != _currentFile)
                {
                    _currentFile = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsStopped => !IsRun;

        public bool IsRun
        {
            get { return _isRun; }
            set
            {
                if (value != _isRun)
                {
                    _isRun = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsStopped));
                }
            }
        }

        public string SourcePath
        {
            get { return _sourcePath; }
            set
            {
                if (value != _sourcePath)
                {
                    _sourcePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string DestinationPath
        {
            get { return _destinationPath; }
            set
            {
                if (value != _destinationPath)
                {
                    _destinationPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public void SelectSourcePath() => SourcePath = SelectPath("Select source folder (to MSG's profile)", SourcePath);

        public void SelectDestinationPath() => DestinationPath = SelectPath("Select destination folder", DestinationPath);

        public ViewModel()
        {
            AutoDestinationPath();
        }
        
        private static string SelectPath(string title, string currentPath = null)
        {
            using (
                var dialog = new CommonOpenFileDialog
                {
                    IsFolderPicker = true,
                    EnsureReadOnly = true,
                    EnsurePathExists = true,
                    InitialDirectory = currentPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)),
                    AllowNonFileSystemItems = false,
                    Multiselect = false,
                    Title = title
                })
            {
                return dialog.ShowDialog() == CommonFileDialogResult.Ok ? dialog.FileName : "";
            }
        }

        public void Run()
        {
            ThreadStart starter = ConvertStore;
            starter += () =>
            {
                ThreadCallback(_message);
            };
            _thread = new Thread(starter) { Name = "Main Convert Thread", IsBackground = true };
            _start = DateTime.Now;
            IsRun = true;
            _thread.Start();
            _tasks.Clear();
        }

        public void Stop()
        {
            if (_thread?.IsAlive == true)
            {
                ThreadCallback("Aborted");
                _thread.Abort();
            }
            else
                ThreadCallback();
        }

        private void ThreadCallback(string message = null)
        {
            var finish = DateTime.Now;
            _tasks.Clear();
            if (!string.IsNullOrWhiteSpace(message))
                MessageBox.Show(message, "Conversion", MessageBoxButton.OK, MessageBoxImage.Information);
            MessageBox.Show((finish - _start).ToString());
            IsRun = false;
            CurrentFile = string.Empty;
        }

        private void ConvertStore()
        {
            if (Process.GetProcessesByName("OUTLOOK").Any())
            {
                MessageBox.Show("Outlook is running. Please close program and rerun conversion.", "Critical error", MessageBoxButton.OK, MessageBoxImage.Stop);
                _message = null;
                return;
            }
            var name = Path.GetFileName(SourcePath);
            var pstPath = Path.Combine(DestinationPath, name + ".pst");
            {
                if (MessageBox.Show($"File {pstPath} already exist. Rewrite?\nAttention! You're lost all of your previous data!", "Rewrite?", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                    return;
                File.Delete(pstPath);
            }
            _session.LogonPstStore(pstPath, 2, name);
            ConvertFolder(SourcePath, _session.Stores.DefaultStore.IPMRootFolder);
            _session.Logoff();
        }

        private void ConvertFolder(string path, IRDOFolder rootFolder)
        {
            if (rootFolder == null)
                return;
            foreach (var directory in Directory.GetDirectories(path))
            {
                var folder = rootFolder.Folders.Add(Path.GetFileName(directory));
                if (folder != null)
                    foreach (var file in Directory.GetFiles(directory))
                    {
                        CurrentFile = file;
                        var message = folder.Items.Add();
                        message.Import(file, 3);
                        message.Sent = true;
                        message.Save();
                    }
            }
        }

        private void AutoDestinationPath()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Pliki programu Outlook");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            DestinationPath = path;
        }

        #region OnPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}