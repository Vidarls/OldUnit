using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OldUnit.Approvals.InteropWrapper
{
    public class FilePathFinder : IFilePathFinder
    {
        public string Find(string fileName)
        {
            var currentDir = new DirectoryInfo(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)));
            var result = SearchDirectory(currentDir, fileName);
            if (result.ToLower() != fileName.ToLower())
                return result;
    
            result = SearchSubDirectories(currentDir, fileName);
            if (result.ToLower() != fileName.ToLower())
                return result;
            
            result = SearchDirectory(currentDir.Parent, fileName);
            if (result.ToLower() != fileName.ToLower())
                return result;
            
            result = SearchSubDirectories(currentDir.Parent, fileName);
            if (result.ToLower() != fileName.ToLower())
                return result;

            return fileName;
        }

        string SearchSubDirectories(DirectoryInfo currentDir, string fileName)
        {
            foreach (var subDir in currentDir.EnumerateDirectories())
            {
                var result = SearchDirectory(subDir, fileName);
                if (result.ToLower() != fileName.ToLower())
                    return result;
            }
            return fileName;
        }

        string SearchDirectory(DirectoryInfo currentDir, string fileName)
        {
            foreach (var file in currentDir.EnumerateFiles())
            {
                if (file.Name.ToLower() == fileName.ToLower())
                    return file.FullName;
            }

            return fileName;
        }
    }
}
