namespace mpDwgBase.Windows
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Threading;
    using Models;
    using ModPlusAPI.Windows;
    using ModPlusService.Dto;
    using ModPlusStyle.Controls.Dialogs;
    using Utils;
    using MessageBox = ModPlusAPI.Windows.MessageBox;

    public partial class BaseUploading
    {
        private const string LangItem = "mpDwgBase";

        // Путь к папке с базами dwg плагина
        private readonly string _dwgBaseFolder;

        // Список значений 
        private readonly List<DwgBaseItem> _dwgBaseItems;
        private List<FileToBind> _filesToBind;

        // Create a Delegate that matches the Signature of the ProgressBar's SetValue method
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);

        private delegate void UpdateProgressTextDelegate(DependencyProperty dp, object value);

        // current archive to upload
        private string _currentFileToUpload;

        public BaseUploading(string dwgBaseFolder, List<DwgBaseItem> dwgBaseItems)
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h26");
            _dwgBaseFolder = dwgBaseFolder;
            _dwgBaseItems = dwgBaseItems;
        }

        private void BaseUploading_OnLoaded(object sender, RoutedEventArgs e)
        {
            // button visibility
            BtSeeArchive.Visibility = Visibility.Collapsed;
            BtUploadArchive.Visibility = Visibility.Collapsed;
            BtDeleteArchive.Visibility = Visibility.Collapsed;

            _filesToBind = new List<FileToBind>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                var file = Path.Combine(_dwgBaseFolder, dwgBaseItem.SourceFile);
                if (File.Exists(file))
                {
                    var fi = new FileInfo(file);
                    var fileToBind = new FileToBind
                    {
                        FileName = fi.Name,
                        FullFileName = fi.FullName,
                        SourceFile = dwgBaseItem.SourceFile,
                        Selected = false,
                        FullDirectory = fi.DirectoryName,
                        SubDirectory = fi.DirectoryName?.Replace($@"{_dwgBaseFolder}\", string.Empty)
                    };
                    if (!HasFileToBindInList(fileToBind))
                    {
                        _filesToBind.Add(fileToBind);
                    }
                }
            }

            LvDwgFiles.ItemsSource = _filesToBind;
        }

        private bool HasFileToBindInList(FileToBind fileToBind)
        {
            var has = false;
            foreach (var toBind in _filesToBind)
            {
                if (toBind.FileName.Equals(fileToBind.FileName) &
                    toBind.FullFileName.Equals(fileToBind.FullFileName) &
                    toBind.SourceFile.Equals(fileToBind.SourceFile))
                {
                    has = true;
                    break;
                }
            }

            return has;
        }

        private void LvDwgFiles_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lv = sender as ListView;
            if (!(lv?.SelectedItem is FileToBind selectedFile))
            {
                return;
            }

            LvItemsInFile.ItemsSource = null;
            var lstToBind = new List<ItemToBind>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (dwgBaseItem.SourceFile.Equals(selectedFile.SourceFile))
                {
                    var itemToBnd = new ItemToBind
                    {
                        Name = dwgBaseItem.Name,
                        Description = dwgBaseItem.Description,
                        Author = dwgBaseItem.Author,
                        Source = dwgBaseItem.Source
                    };
                    if (dwgBaseItem.IsBlock)
                    {
                        itemToBnd.IsBlock = Visibility.Visible;
                        itemToBnd.IsDrawing = Visibility.Collapsed;
                    }
                    else
                    {
                        itemToBnd.IsBlock = Visibility.Collapsed;
                        itemToBnd.IsDrawing = Visibility.Visible;
                    }

                    if (!lstToBind.Contains(itemToBnd))
                    {
                        lstToBind.Add(itemToBnd);
                    }
                }
            }

            LvItemsInFile.ItemsSource = lstToBind;
        }

        private void BtMakeArchive_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TbFeedback.Text))
            {
                if (!MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "h75")))
                {
                    return;
                }
            }

            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
            if (!_filesToBind.Any(x => x.Selected))
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg15"), MessageBoxIcon.Alert);
                return;
            }

            CreateArchive();
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, string.Empty);
            Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, 0.0);
        }

        private void CreateArchive()
        {
            // Create a new instance of our ProgressBar Delegate that points
            //  to the ProgressBar's SetValue method.
            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

            // progress text
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, ModPlusAPI.Language.GetItem(LangItem, "msg16"));

            // create temp folder
            var tmpFolder = Path.Combine(_dwgBaseFolder, "Temp");
            if (!Directory.Exists(tmpFolder))
            {
                Directory.CreateDirectory(tmpFolder);
            }

            // create base file with selected files items
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = _dwgBaseItems.Count;
            ProgressBar.Value = 0;

            // Соберем список файлов-источников чтобы проще их сверять
            var sourceFiles = new List<string>();
            foreach (var fileToBind in _filesToBind)
            {
                if (!sourceFiles.Contains(fileToBind.SourceFile))
                {
                    sourceFiles.Add(fileToBind.SourceFile);
                }
            }

            var baseFileToArchive = new List<DwgBaseItem>();
            for (var i = 0; i < _dwgBaseItems.Count; i++)
            {
                Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "msg17")}: {i}/{_dwgBaseItems.Count}");
                Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)i);
                var dwgBaseItem = _dwgBaseItems[i];
                if (sourceFiles.Contains(dwgBaseItem.SourceFile))
                {
                    if (!baseFileToArchive.Contains(dwgBaseItem))
                    {
                        baseFileToArchive.Add(dwgBaseItem);
                    }
                }
            }

            // save xml file
            var xmlToArchive = Path.Combine(tmpFolder, "UserDwgBase.xml");
            DwgBaseHelpers.SerializerToXml(baseFileToArchive, xmlToArchive);
            if (!File.Exists(xmlToArchive))
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg18"), MessageBoxIcon.Close);
                return;
            }

            // comment file
            var commentFile = CreateCommentFile(tmpFolder);

            // create zip
            _currentFileToUpload = Path.ChangeExtension(Path.Combine(tmpFolder, Path.GetRandomFileName()), ".zip");
            if (File.Exists(_currentFileToUpload))
            {
                File.Delete(_currentFileToUpload);
            }

            using (var zip = ZipFile.Open(_currentFileToUpload, ZipArchiveMode.Create))
            {
                // create directories
                foreach (var fileToBind in _filesToBind.Where(x => x.Selected))
                {
                    zip.CreateEntryFromFile(
                        fileToBind.FullFileName,
                        Path.Combine(fileToBind.SubDirectory, fileToBind.FileName));
                }

                // add xml file and delete him
                zip.CreateEntryFromFile(xmlToArchive, new FileInfo(xmlToArchive).Name);

                // add comment file
                if (!string.IsNullOrEmpty(commentFile) && File.Exists(commentFile))
                {
                    zip.CreateEntryFromFile(commentFile, new FileInfo(commentFile).Name);
                }
            }

            File.Delete(xmlToArchive);

            // show buttons
            BtMakeArchive.Visibility = Visibility.Collapsed;
            BtSeeArchive.Visibility = Visibility.Visible;
            BtDeleteArchive.Visibility = Visibility.Visible;
            BtUploadArchive.Visibility = Visibility.Visible;
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, ModPlusAPI.Language.GetItem(LangItem, "msg19"));
        }

        // window closed
        private void BaseUploading_OnClosed(object sender, EventArgs e)
        {
            var tmpFolder = Path.Combine(_dwgBaseFolder, "Temp");
            try
            {
                // check if temp directory exist
                if (Directory.Exists(tmpFolder))
                {
                    Directory.Delete(tmpFolder, true);
                }
            }
            catch
            {
                MessageBox.Show(
                    $"{ModPlusAPI.Language.GetItem(LangItem, "msg20")}: {tmpFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg21")}");
            }
        }

        // open
        private void BtSeeArchive_OnClick(object sender, RoutedEventArgs e)
        {
            if (File.Exists(_currentFileToUpload))
            {
                Process.Start(_currentFileToUpload);
            }
            else
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg22"));
            }
        }

        // upload
        private async void BtUploadArchive_OnClick(object sender, RoutedEventArgs e)
        {
            if (!await ModPlusAPI.Web.Connection.HasAllConnectionAsync(1))
            {
                await this.ShowMessageAsync(ModPlusAPI.Language.GetItem(LangItem, "msg23"), string.Empty);
                return;
            }

            if (File.Exists(_currentFileToUpload))
            {
                ProgressDialogController controller = null;
                try
                {
                    var settings = new MetroDialogSettings
                    {
                        AnimateShow = true,
                        AnimateHide = true,
                        DialogTitleFontSize = 20
                    };
                    controller = await this.ShowProgressAsync(ModPlusAPI.Language.GetItem(LangItem, "msg24"), string.Empty, false, settings);
                    controller.Minimum = 0;
                    controller.Maximum = 2;

                    using (var client = ModPlusAPI.Web.ApiClient.CreateClient())
                    {
                        controller.SetMessage(ModPlusAPI.Language.GetItem(LangItem, "msg25"));
                        controller.SetProgress(1);
                        await client.UploadUserFile(new[] { "DwgForBaseFromUsers" }, _currentFileToUpload, false, true);

                        controller.SetMessage(ModPlusAPI.Language.GetItem(LangItem, "msg26"));
                        controller.SetProgress(2);
                        var emailSettings = await client.GetEmailSettings();
                        await SendEmailNotification(emailSettings);
                    }

                    await controller.CloseAsync();

                    await this.ShowMessageAsync(
                        string.Empty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "msg27")}: {_currentFileToUpload} {ModPlusAPI.Language.GetItem(LangItem, "msg28")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg29")}");
                }
                catch (Exception exception)
                {
                    if (controller != null)
                        await controller.CloseAsync();
                    ExceptionBox.Show(exception);
                }
            }
            else
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg22"));
            }
        }

        // create comment file
        private string CreateCommentFile(string tmpFolder)
        {
            var file = string.Empty;
            if (!string.IsNullOrEmpty(TbComment.Text))
            {
                file = Path.Combine(tmpFolder, "comment.txt");
                var text = $"Feedback :{TbFeedback.Text}{Environment.NewLine}{TbComment.Text}";
                File.WriteAllText(file, text);
            }

            return file;
        }

        private void BtDeleteArchive_OnClick(object sender, RoutedEventArgs e)
        {
            if (File.Exists(_currentFileToUpload))
            {
                if (MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "msg30"), MessageBoxIcon.Question))
                {
                    var wasDel = false;
                    try
                    {
                        File.Delete(_currentFileToUpload);
                        wasDel = true;
                    }
                    catch
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg31"));
                    }

                    if (wasDel)
                    {
                        BtDeleteArchive.Visibility = Visibility.Collapsed;
                        BtUploadArchive.Visibility = Visibility.Collapsed;
                        BtSeeArchive.Visibility = Visibility.Collapsed;
                        BtMakeArchive.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private async Task SendEmailNotification(EmailSettingsDto emailSettings)
        {
            if (emailSettings != null && emailSettings.IsAvailable())
            {
                var from = new MailAddress(emailSettings.Address, "Dwg Base Notification");
                var to = new MailAddress(emailSettings.Address);
                using (var m = new MailMessage(from, to))
                {
                    m.Subject = "Dwg Base Notification";
                    m.Body = "New file was uploaded by user";
                    m.IsBodyHtml = false;
                    using (var smtp = new SmtpClient(emailSettings.Host, emailSettings.Port))
                    {
                        smtp.Credentials = new NetworkCredential(emailSettings.UserName, emailSettings.Password);
                        smtp.EnableSsl = false;
                        await smtp.SendMailAsync(m).ConfigureAwait(true);
                    }
                }
            }
        }
    }
}
