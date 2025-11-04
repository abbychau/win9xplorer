namespace win9xplorer
{
    /// <summary>
    /// Utility methods for the file explorer
    /// </summary>
    internal static class ExplorerUtils
    {
        /// <summary>
        /// Formats file size in a human-readable format
        /// </summary>
        public static string FormatFileSize(long bytes)
        {
            if (bytes == 0) return "0 bytes";
            
            string[] sizes = { "bytes", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len /= 1024;
            }
            
            if (order == 0)
                return $"{len:0} {sizes[order]}";
            else
                return $"{len:0.#} {sizes[order]}";
        }

        /// <summary>
        /// Gets a user-friendly file type description
        /// </summary>
        public static string GetFileType(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                return "File";
                
            // Use Windows API to get the actual file type description
            WinApi.SHFILEINFO shfi = new WinApi.SHFILEINFO();
            uint flags = WinApi.SHGFI_TYPENAME | WinApi.SHGFI_USEFILEATTRIBUTES;
            
            WinApi.SHGetFileInfo(extension, 0, ref shfi, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shfi), flags);
            
            if (!string.IsNullOrEmpty(shfi.szTypeName))
            {
                return shfi.szTypeName;
            }
            
            // Fallback to custom descriptions
            return extension.ToUpper() switch
            {
                ".TXT" => "Text Document",
                ".DOC" or ".DOCX" => "Microsoft Word Document",
                ".XLS" or ".XLSX" => "Microsoft Excel Worksheet",
                ".PDF" => "Adobe Acrobat Document",
                ".EXE" => "Application",
                ".DLL" => "Application Extension",
                ".ZIP" => "Compressed (zipped) Folder",
                ".JPG" or ".JPEG" => "JPEG Image",
                ".PNG" => "PNG Image",
                ".GIF" => "GIF Image",
                ".BMP" => "Bitmap Image",
                ".MP3" => "MP3 Format Sound",
                ".MP4" => "MP4 Video",
                ".AVI" => "Video Clip",
                ".HTM" or ".HTML" => "Internet Document",
                _ => extension.Length > 0 ? $"{extension.ToUpper().Substring(1)} File" : "File"
            };
        }

        /// <summary>
        /// Formats free space with percentage
        /// </summary>
        public static string FormatFreeSpace(long totalBytes, long freeBytes)
        {
            if (totalBytes == 0) return "0 bytes";
            
            double freePercentage = ((double)freeBytes / totalBytes) * 100;
            string sizeText = FormatFileSize(freeBytes);
            
            return $"{sizeText} ({freePercentage:0.#}% free)";
        }

        /// <summary>
        /// Checks if a path is a drive root (e.g., "C:\")
        /// </summary>
        public static bool IsDriveRoot(string path)
        {
            return !string.IsNullOrEmpty(path) && Path.GetPathRoot(path) == path;
        }

        /// <summary>
        /// Gets the parent directory of a path, handling drive roots appropriately
        /// </summary>
        public static string? GetParentPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return null;

            var parent = Directory.GetParent(path);
            if (parent != null)
            {
                return parent.FullName;
            }
            else if (IsDriveRoot(path))
            {
                // If we're at a drive root (like C:\), return empty string to indicate My Computer
                return "";
            }

            return null;
        }

        /// <summary>
        /// Safely gets drive information
        /// </summary>
        public static (string label, string totalSize, string freeSpace, string usedSpace, string fileSystem) GetDriveInfo(DriveInfo drive)
        {
            string driveLabel = "";
            string totalSize = "";
            string freeSpace = "";
            string usedSpace = "";
            string fileSystem = "";
            
            if (drive.IsReady)
            {
                driveLabel = string.IsNullOrEmpty(drive.VolumeLabel) ? 
                    $"{drive.Name.TrimEnd('\\')}" : 
                    $"{drive.VolumeLabel} ({drive.Name.TrimEnd('\\')})";
                
                try
                {
                    // Get drive space information
                    long totalBytes = drive.TotalSize;
                    long freeBytes = drive.AvailableFreeSpace;
                    long usedBytes = totalBytes - freeBytes;
                    
                    totalSize = FormatFileSize(totalBytes);
                    freeSpace = FormatFreeSpace(totalBytes, freeBytes);
                    usedSpace = FormatFileSize(usedBytes);
                    fileSystem = drive.DriveFormat;
                }
                catch
                {
                    totalSize = "Unknown";
                    freeSpace = "Unknown";
                    usedSpace = "Unknown";
                    fileSystem = "Unknown";
                }
            }
            else
            {
                driveLabel = $"{drive.Name.TrimEnd('\\')} ({drive.DriveType})";
            }

            return (driveLabel, totalSize, freeSpace, usedSpace, fileSystem);
        }
    }
}