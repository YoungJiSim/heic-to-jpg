using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Explorers _explorers;
        private Outlook.Explorer _activeExplorer;
        private Outlook.MailItem _currentMail = null;
        private readonly string logFilePath = Path.Combine(Path.GetTempPath(), "outlook_heic_log.txt");

        private void WriteToFileLog(string message)
        {
            try
            {
                string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                File.AppendAllText(logFilePath, $"[{timeStamp}] {message}\r\n");
            }
            catch { }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (File.Exists(logFilePath)) File.Delete(logFilePath);

                Debug.WriteLine("[[Log]] Initializing Outlook Watcher...");
                WriteToFileLog("[[Log]] Initializing Outlook Watcher...");

                _explorers = this.Application.Explorers;
                _explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(Explorers_NewExplorer);

                _activeExplorer = this.Application.ActiveExplorer();
                if (_activeExplorer != null)
                {
                    _activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ActiveExplorer_SelectionChange);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                WriteToFileLog($"[[Error]] Startup failed: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnhookCurrentMail();
            _explorers = null;
            _activeExplorer = null;
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                _activeExplorer = explorer;
                _activeExplorer.SelectionChange -= ActiveExplorer_SelectionChange;
                _activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ActiveExplorer_SelectionChange);
            }
            catch { }
        }

        private void ActiveExplorer_SelectionChange()
        {
            try
            {
                UnhookCurrentMail();

                if (_activeExplorer.Selection != null && _activeExplorer.Selection.Count > 0)
                {
                    object selectedItem = _activeExplorer.Selection[1];
                    if (selectedItem is Outlook.MailItem)
                    {
                        _currentMail = (Outlook.MailItem)selectedItem;
                        _currentMail.BeforeAttachmentRead += new Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler(CurrentMail_BeforeAttachmentRead);
                    }
                    else
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(selectedItem);
                    }
                }
            }
            catch { }
        }

        private void CurrentMail_BeforeAttachmentRead(Outlook.Attachment Attachment, ref bool Cancel)
        {
            try
            {
                if (Attachment == null) return;

                string fileName = Attachment.FileName;

                if (!fileName.EndsWith(".heic", StringComparison.OrdinalIgnoreCase)) return;

                Cancel = true;

                Debug.WriteLine($"[[Log]] Target HEIC Intercepted Safely: {fileName}");
                WriteToFileLog($"[[Log]] Target HEIC Intercepted Safely: {fileName}");

                _ = OpenHeicInExternalViewerAsync(Attachment, fileName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[[Error]] Main Attachment Intercept Failed: {ex.Message}");
                WriteToFileLog($"[[Error]] Main Attachment Intercept Failed: {ex.Message}");
            }
        }

        private async Task OpenHeicInExternalViewerAsync(Outlook.Attachment attachment, string fileName)
        {
            try
            {
                string tempFolder = Path.GetTempPath();

                string mailSubject = _currentMail != null ? _currentMail.Subject : "NoSubject";
                string receivedDateStr = _currentMail != null ? _currentMail.ReceivedTime.ToString("yyyyMMdd_HHmmss") : "NoDate";

                string safeSubject = MakeSafeFileName(mailSubject);
                string safeFileName = MakeSafeFileName(Path.GetFileNameWithoutExtension(fileName));

                string localHeicPath = Path.Combine(tempFolder, fileName);
                string cacheJpgName = $"{receivedDateStr}_{safeSubject}_{safeFileName}.jpg";
                string localJpgPath = Path.Combine(tempFolder, cacheJpgName);

                Debug.WriteLine($"[[Log]] Target info parsed. Expected cache name: {cacheJpgName}");
                WriteToFileLog($"[[Log]] Target info parsed. Expected cache name: {cacheJpgName}");

                if (File.Exists(localJpgPath))
                {
                    Process.Start(localJpgPath);
                    Debug.WriteLine($"[[Cache Hit]] Existing JPG found. Launching external viewer: {cacheJpgName}");
                    WriteToFileLog($"[[Cache Hit]] Existing JPG found. Launching external viewer: {cacheJpgName}");
                    return;
                }

                Debug.WriteLine("[[Cache Miss]] Launching Python conversion pipeline...");
                WriteToFileLog("[[Cache Miss]] Launching Python conversion pipeline...");

                attachment.SaveAsFile(localHeicPath);
                Debug.WriteLine($"[[Log]] HEIC file temporarily saved to: {localHeicPath}");
                WriteToFileLog($"[[Log]] HEIC file temporarily saved to: {localHeicPath}");

                string pythonExePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "heic_to_jpg.exe");

                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = pythonExePath,
                    Arguments = $"\"{localHeicPath}\" \"{localJpgPath}\"",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    UseShellExecute = false
                };

                using (Process process = Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        await Task.Run(() => process.WaitForExit());
                        Debug.WriteLine($"[[Log]] Python process exited with code: {process.ExitCode}");
                        WriteToFileLog($"[[Log]] Python process exited with code: {process.ExitCode}");
                    }
                }

                if (File.Exists(localHeicPath))
                {
                    File.Delete(localHeicPath);
                    Debug.WriteLine($"[[Log]] Temporary HEIC file deleted.");
                    WriteToFileLog($"[[Log]] Temporary HEIC file deleted.");
                }

                if (File.Exists(localJpgPath))
                {
                    Process.Start(localJpgPath);
                    Debug.WriteLine($"[[Success]] Converted & Opened in external viewer: {cacheJpgName}");
                    WriteToFileLog($"[[Success]] Converted & Opened in external viewer: {cacheJpgName}");
                }
                else
                {
                    Debug.WriteLine($"[[Error]] Expected output JPG not found after Python execution: {localJpgPath}");
                    WriteToFileLog($"[[Error]] Expected output JPG not found after Python execution: {localJpgPath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[[Error]] Async external viewer generator failed: {ex.Message}");
                WriteToFileLog($"[[Error]] Async external viewer generator failed: {ex.Message}");
            }
        }

        private string MakeSafeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "noname";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name.Replace(" ", "_").Trim();
        }

        private void UnhookCurrentMail()
        {
            try
            {
                if (_currentMail != null)
                {
                    _currentMail.BeforeAttachmentRead -= CurrentMail_BeforeAttachmentRead;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_currentMail);
                    _currentMail = null;
                }
            }
            catch { }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}