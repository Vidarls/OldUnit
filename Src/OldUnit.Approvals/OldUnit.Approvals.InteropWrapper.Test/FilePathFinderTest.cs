using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Xunit;

namespace OldUnit.Approvals.InteropWrapper.Test
{
    public class FilePathFinderTest
    {
        private const string SubDirName = "TestSubDir";
        private const string FileInSubdirName = "TestFile.test";

        private void CreateTestSubDir()
        {
            var testSubDirPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)),SubDirName);
            if (!Directory.Exists(testSubDirPath))
                Directory.CreateDirectory(testSubDirPath);
            var testFilePath = Path.Combine(testSubDirPath, FileInSubdirName);
            if (!File.Exists(testFilePath))
                File.CreateText(testFilePath).WriteLine("Test");

            
        }

        [Fact]
        [ApprovalTests.Reporters.UseReporter(typeof(ApprovalTests.Reporters.FileLauncherReporter))]
        public void Can_find_file_in_same_folder_as_executing()
        {
            IFilePathFinder finder = new FilePathFinder();
            var path = finder.Find("ApprovalTests.dll");
            ApprovalTests.Approvals.Approve(path);
        }

        [Fact]
        [ApprovalTests.Reporters.UseReporter(typeof(ApprovalTests.Reporters.FileLauncherReporter))]
        public void Can_find_file_in_subfolde_one_level_below()
        {
            IFilePathFinder finder = new FilePathFinder();
            CreateTestSubDir();
            var path = finder.Find(FileInSubdirName);
            ApprovalTests.Approvals.Approve(path);
        }

        [Fact]
        [ApprovalTests.Reporters.UseReporter(typeof(ApprovalTests.Reporters.FileLauncherReporter))]
        public void Can_find_file_in_parent_folder()
        {
            IFilePathFinder finder = new FilePathFinder();
            var path = finder.Find(".gitignore");
            ApprovalTests.Approvals.Approve(path);
        }

        [Fact]
        [ApprovalTests.Reporters.UseReporter(typeof(ApprovalTests.Reporters.FileLauncherReporter))]
        public void Can_find_file_in_paralell_folder()
        {
            IFilePathFinder finder = new FilePathFinder();
            var path = finder.Find("TortoiseIDiff.exe");
            ApprovalTests.Approvals.Approve(path);
        }

    }
}
