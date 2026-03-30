using System.Collections.Specialized;
using Microsoft.VisualBasic.FileIO;

namespace win9xplorer
{
    internal enum FileConflictStrategy
    {
        AskUser = 0,
        OverwriteExisting = 1,
        SkipExisting = 2
    }

    internal sealed class FileOperationResult
    {
        public int SuccessCount { get; set; }
        public int SkippedCount { get; set; }
        public bool IsCanceled { get; set; }
        public List<string> Errors { get; } = new();
    }

    internal sealed class FileOperationsService
    {
        private const string PreferredDropEffect = "Preferred DropEffect";
        private const int DropEffectMove = 2;

        public void CopyToClipboard(IReadOnlyCollection<string> filePaths)
        {
            if (filePaths.Count == 0)
                return;

            var files = new StringCollection();
            files.AddRange(filePaths.ToArray());
            Clipboard.SetFileDropList(files);
        }

        public void CutToClipboard(IReadOnlyCollection<string> filePaths)
        {
            if (filePaths.Count == 0)
                return;

            var files = new StringCollection();
            files.AddRange(filePaths.ToArray());

            var dataObject = new DataObject();
            dataObject.SetFileDropList(files);
            dataObject.SetData(PreferredDropEffect, BitConverter.GetBytes(DropEffectMove));
            Clipboard.SetDataObject(dataObject, true);
        }

        public bool TryGetClipboardFileDrop(out List<string> filePaths, out bool isCutOperation)
        {
            filePaths = new List<string>();
            isCutOperation = false;

            if (!Clipboard.ContainsFileDropList())
                return false;

            foreach (string? file in Clipboard.GetFileDropList())
            {
                if (!string.IsNullOrWhiteSpace(file))
                {
                    filePaths.Add(file);
                }
            }

            try
            {
                var dataObject = Clipboard.GetDataObject();
                if (dataObject?.GetDataPresent(PreferredDropEffect) == true
                    && dataObject.GetData(PreferredDropEffect) is byte[] dropEffect
                    && dropEffect.Length >= sizeof(int))
                {
                    isCutOperation = BitConverter.ToInt32(dropEffect, 0) == DropEffectMove;
                }
            }
            catch
            {
            }

            return filePaths.Count > 0;
        }

        public FileOperationResult PasteToDirectory(
            IEnumerable<string> sourcePaths,
            string targetDirectory,
            bool isMove,
            FileConflictStrategy conflictStrategy = FileConflictStrategy.AskUser)
        {
            var result = new FileOperationResult();

            foreach (string filePath in sourcePaths)
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    continue;

                try
                {
                    string fileName = Path.GetFileName(filePath);
                    if (string.IsNullOrWhiteSpace(fileName))
                        continue;

                    string destinationPath = Path.Combine(targetDirectory, fileName);

                    if (string.Equals(filePath, destinationPath, StringComparison.OrdinalIgnoreCase))
                    {
                        result.SkippedCount++;
                        continue;
                    }

                    if (Directory.Exists(filePath))
                    {
                        if (conflictStrategy == FileConflictStrategy.AskUser)
                        {
                            if (isMove)
                                FileSystem.MoveDirectory(filePath, destinationPath, UIOption.AllDialogs, UICancelOption.ThrowException);
                            else
                                FileSystem.CopyDirectory(filePath, destinationPath, UIOption.AllDialogs, UICancelOption.ThrowException);

                            result.SuccessCount++;
                        }
                        else
                        {
                            HandleDirectoryOperation(filePath, destinationPath, isMove, conflictStrategy, result);
                        }
                    }
                    else if (File.Exists(filePath))
                    {
                        if (conflictStrategy == FileConflictStrategy.AskUser)
                        {
                            if (isMove)
                                FileSystem.MoveFile(filePath, destinationPath, UIOption.AllDialogs, UICancelOption.ThrowException);
                            else
                                FileSystem.CopyFile(filePath, destinationPath, UIOption.AllDialogs, UICancelOption.ThrowException);

                            result.SuccessCount++;
                        }
                        else
                        {
                            HandleFileOperation(filePath, destinationPath, isMove, conflictStrategy, result);
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    result.IsCanceled = true;
                    break;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"'{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            return result;
        }

        public Task<FileOperationResult> PasteToDirectoryAsync(
            IEnumerable<string> sourcePaths,
            string targetDirectory,
            bool isMove,
            FileConflictStrategy conflictStrategy = FileConflictStrategy.AskUser)
        {
            return RunOnStaThread(() => PasteToDirectory(sourcePaths, targetDirectory, isMove, conflictStrategy));
        }

        private static void HandleFileOperation(
            string sourcePath,
            string destinationPath,
            bool isMove,
            FileConflictStrategy conflictStrategy,
            FileOperationResult result)
        {
            bool destinationExists = File.Exists(destinationPath);
            if (destinationExists && conflictStrategy == FileConflictStrategy.SkipExisting)
            {
                result.SkippedCount++;
                return;
            }

            if (isMove)
            {
                if (destinationExists && conflictStrategy == FileConflictStrategy.OverwriteExisting)
                {
                    File.Delete(destinationPath);
                }

                File.Move(sourcePath, destinationPath);
            }
            else
            {
                File.Copy(sourcePath, destinationPath, overwrite: conflictStrategy == FileConflictStrategy.OverwriteExisting);
            }

            result.SuccessCount++;
        }

        private static void HandleDirectoryOperation(
            string sourcePath,
            string destinationPath,
            bool isMove,
            FileConflictStrategy conflictStrategy,
            FileOperationResult result)
        {
            if (IsSubPath(sourcePath, destinationPath))
            {
                throw new IOException("Cannot move or copy a folder into itself.");
            }

            bool destinationExists = Directory.Exists(destinationPath);
            if (destinationExists && conflictStrategy == FileConflictStrategy.SkipExisting)
            {
                result.SkippedCount++;
                return;
            }

            if (!destinationExists)
            {
                if (isMove)
                {
                    Directory.Move(sourcePath, destinationPath);
                }
                else
                {
                    CopyDirectoryContents(sourcePath, destinationPath, overwriteFiles: true);
                }

                result.SuccessCount++;
                return;
            }

            // destination exists and strategy is overwrite
            CopyDirectoryContents(sourcePath, destinationPath, overwriteFiles: true);
            if (isMove)
            {
                Directory.Delete(sourcePath, true);
            }

            result.SuccessCount++;
        }

        private static void CopyDirectoryContents(string sourcePath, string destinationPath, bool overwriteFiles)
        {
            Directory.CreateDirectory(destinationPath);

            foreach (string filePath in Directory.EnumerateFiles(sourcePath))
            {
                string fileName = Path.GetFileName(filePath);
                string destinationFilePath = Path.Combine(destinationPath, fileName);
                File.Copy(filePath, destinationFilePath, overwriteFiles);
            }

            foreach (string directoryPath in Directory.EnumerateDirectories(sourcePath))
            {
                string directoryName = Path.GetFileName(directoryPath);
                string destinationDirectoryPath = Path.Combine(destinationPath, directoryName);
                CopyDirectoryContents(directoryPath, destinationDirectoryPath, overwriteFiles);
            }
        }

        private static bool IsSubPath(string parentPath, string candidatePath)
        {
            string normalizedParent = Path.GetFullPath(parentPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            string normalizedCandidate = Path.GetFullPath(candidatePath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            return normalizedCandidate.StartsWith(normalizedParent, StringComparison.OrdinalIgnoreCase);
        }

        public FileOperationResult DeletePaths(IEnumerable<string> filePaths)
        {
            var result = new FileOperationResult();

            foreach (string filePath in filePaths)
            {
                try
                {
                    if (Directory.Exists(filePath))
                    {
                        FileSystem.DeleteDirectory(
                            filePath,
                            UIOption.OnlyErrorDialogs,
                            RecycleOption.SendToRecycleBin,
                            UICancelOption.ThrowException);
                        result.SuccessCount++;
                    }
                    else if (File.Exists(filePath))
                    {
                        FileSystem.DeleteFile(
                            filePath,
                            UIOption.OnlyErrorDialogs,
                            RecycleOption.SendToRecycleBin,
                            UICancelOption.ThrowException);
                        result.SuccessCount++;
                    }
                }
                catch (OperationCanceledException)
                {
                    result.IsCanceled = true;
                    break;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"'{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            return result;
        }

        public Task<FileOperationResult> DeletePathsAsync(IEnumerable<string> filePaths)
        {
            return RunOnStaThread(() => DeletePaths(filePaths));
        }

        private static Task<T> RunOnStaThread<T>(Func<T> work)
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

            var thread = new Thread(() =>
            {
                try
                {
                    T result = work();
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }
    }
}
