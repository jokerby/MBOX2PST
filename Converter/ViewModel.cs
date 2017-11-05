using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Converter.Add;
using Converter.Annotations;
using Redemption;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Converter
{
    public sealed class ViewModel : INotifyPropertyChanged
    {
        private readonly RDOSession _session = new RDOSession();
        private string _destinationPath;
        private string _sourcePath;
        private Thread _thread;
        private DateTime _start;
        private readonly List<Task> _tasks = new List<Task>();
        private string _message = "Finished";
        private string _currentFile = string.Empty;
        private bool _isRun;

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

        public ViewModel()
        {
            AutoSourcePath();
            AutoDestinationPath();
        }

        public void SelectSourcePath() => SourcePath = SelectPath("Select source folder (to Thunderbird's profile)", SourcePath);

        public void SelectDestinationPath() => DestinationPath = SelectPath("Select destination folder", DestinationPath);

        public void Run()
        {
            ThreadStart starter = ConvertMail;
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
            _tasks.Clear();
            if (!string.IsNullOrWhiteSpace(message))
            {
                MessageBox.Show(message, "Conversion", MessageBoxButton.OK, MessageBoxImage.Information);
                MessageBox.Show((DateTime.Now - _start).ToString());
            }
            IsRun = false;
            CurrentFile = string.Empty;
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

        private void AutoDestinationPath()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            DestinationPath = path;
        }

        private void AutoSourcePath()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Thunderbird", "Profiles");
            if (Directory.Exists(path))
            {
                var folders = Directory.GetDirectories(path);
                SourcePath = folders.Length == 1
                    ? folders.First()
                    : (folders.Any() ? path : string.Empty);
            }
            else
                SourcePath = string.Empty;
        }

        private void ConvertMail()
        {
            if (Process.GetProcessesByName("OUTLOOK").Any())
            {
                MessageBox.Show("Outlook is running. Please close program and rerun conversion.", "Critical error", MessageBoxButton.OK, MessageBoxImage.Stop);
                _message = null;
                return;
            }
            var source = Path.Combine(SourcePath, "Mail");
            if (!Directory.Exists(source))
                return;
            var psts = Directory.GetDirectories(source);
            foreach (var pst in psts)
                ConvertStore(pst);
        }

        private void ConvertStore(string path)
        {
            var name = Path.GetFileName(path);
            var pstPath = Path.Combine(DestinationPath, name + ".pst");
            if (File.Exists(pstPath))
            {
                if (MessageBox.Show($"File {pstPath} already exist. Rewrite?\nAttention! You're lost all of your previous data!", "Rewrite?", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                    return;
                File.Delete(pstPath);
            }
            _session.LogonPstStore(pstPath, 2, name);
            var root = _session.Stores.DefaultStore.IPMRootFolder;
            ConvertFolder(path, root);
            _session.Logoff();
        }

        private void ConvertFolder(string path, IRDOFolder rootFolder, bool isRoot = true)
        {
            //var factory = new TaskFactory(new LimitedConcurrencyLevelTaskScheduler(4));
            foreach (var file in Directory.GetFiles(path))
                if (!Path.HasExtension(file))
                    ReadFolder(file, GetFolderByNameOrCreate(rootFolder, Path.GetFileNameWithoutExtension(file), isRoot));
            //_tasks.Add(factory.StartNew(() => ReadFolder(file, GetFolderByNameOrCreate(rootFolder, Path.GetFileNameWithoutExtension(file), isRoot))));
            foreach (var directory in Directory.GetDirectories(path))
            {
                var dirName = Path.GetFileName(directory);
                if (dirName?.Tail(4).ToLower() == ".sbd")
                    ConvertFolder(directory, GetFolderByNameOrCreate(rootFolder, dirName.Substring(0, dirName.Length - 4)), false);
            }
            //Task.WaitAll(_tasks.ToArray());
        }

        private IRDOFolder GetFolderByNameOrCreate(IRDOFolder rootFolder, string name, bool isRoot = true)
        {
            if (!isRoot)
                return FindOrCreate(rootFolder, name, 0, false);

            switch (name.ToLower())
            {
                case "inbox":
                    return FindOrCreate(rootFolder, "Skrzynka odbiorcza", (int)rdoDefaultFolders.olFolderInbox);
                case "trash":
                    return FindOrCreate(rootFolder, "Elementy usunięte", (int)rdoDefaultFolders.olFolderDeletedItems);
                case "sent":
                    return FindOrCreate(rootFolder, "Elementy wysłane", (int)rdoDefaultFolders.olFolderSentMail);
                case "junk":
                    return FindOrCreate(rootFolder, "Wiadomości-śmieci", (int)rdoDefaultFolders.olFolderJunk, false);
                case "archives":
                    return FindOrCreate(rootFolder, "Archiwum", (int)rdoDefaultFolders.olFolderArchive, false);
                case "drafts":
                    return FindOrCreate(rootFolder, "Wersje robocze", (int)rdoDefaultFolders.olFolderDrafts);
                case "unsent messages":
                    return FindOrCreate(rootFolder, "Skrzynka nadawcza", (int)rdoDefaultFolders.olFolderOutbox);
                default:
                    return FindOrCreate(rootFolder, name, 0, false);
            }
        }

        private IRDOFolder FindOrCreate(IRDOFolder root, string name, int defKind, bool isDefKindSet = true)
        {
            RDOFolder2 folder;
            try
            {
                if (isDefKindSet)
                    folder = _session.GetDefaultFolder((rdoDefaultFolders)defKind) as RDOFolder2;
                else
                    folder = root.Folders[name] as RDOFolder2;
            }
            catch (Exception e)
            {
                folder = root.Folders.Add(name) as RDOFolder2;
                if (isDefKindSet)
                    folder?.SetAsDefaultFolder(defKind);
            }
            if (folder == null)
            {
                folder = root.Folders.Add(name) as RDOFolder2;
                if (isDefKindSet)
                    folder?.SetAsDefaultFolder(defKind);
            }
            return folder;
        }

        private void ReadFolder(string path, IRDOFolder folder)
        {
            if (folder == null)
                return;
            CurrentFile = path;
            using (var file = new StreamReader(path, Encoding.UTF8))
            {
                var lines = new List<string>();
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    if (line.StartsWith("From - ") && lines.Any())
                    {
                        SaveMessage(lines, folder);
                        lines.Clear();
                    }
                    lines.Add(line);
                }
            }
        }

        private static void SaveMessage([NotNull] IReadOnlyCollection<string> lines, [NotNull] IRDOFolder folder)
        {
            if (!lines.Any())
                return;
            var pathDir = Path.Combine(@"C:\Windows\Temp", folder.Name.TrimIllegalFromPath());
            if (!Directory.Exists(pathDir))
                Directory.CreateDirectory(pathDir);
            var pathMes = Path.Combine(pathDir, lines.First().TrimIllegalFromPath());
            File.WriteAllLines(pathMes, lines);
            var message = folder.Items.Add();
            message.Import(pathMes, 1024);
            message.Sent = folder.Name != "Wersje robocze";
            message.Save();
            File.Delete(pathMes);
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