using IWshRuntimeLibrary;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using File = System.IO.File;
namespace EdaraOffline
{
    public static class Program
    {
        static IConfiguration Config { get; set; }

        static void Main(string[] args)
        {
            Config = new ConfigurationBuilder().AddJsonFile(Path.Combine(Directory.GetCurrentDirectory(), "AppSettings.json")).Build();
            FileInfo chromeExe = GetChromeExe();
            IEnumerable<string> chromeShortcuts = GetChromeShortcutPaths();
            if (args.Length == 0 || args.Any(x => x.ToLower() == "install"))
                InstallOffline(chromeExe, chromeShortcuts);
            else if (args.Any(x => x.ToLower() == "uninstall"))
                UninstallOffline(chromeExe, chromeShortcuts);
            Console.ReadKey();
        }

        public static void InstallOffline(FileInfo chromeExe, IEnumerable<string> chromeShortcuts)
        {
            if (chromeExe != null)
                chromeShortcuts.Where(path => File.Exists(path))
                    .ToList().ForEach(path => SaveShortcustWithArguments(path, chromeExe));
            Console.WriteLine("Installation finished press any key to exit");
        }
        public static void UninstallOffline(FileInfo chromeExe, IEnumerable<string> chromeShortcuts)
        {
            if (IsConnectedToInternet())
            {
                if (chromeExe != null)
                    chromeShortcuts.Where(path => File.Exists(path))
                        .ToList().ForEach(path => SaveShortcustWithoutArguments(path, chromeExe.FullName, chromeExe.FullName));
                Console.WriteLine("Uninstallation finished press any key to exit");
            }
            else
                Console.WriteLine("You must be connected to internet to uninstall");
            Console.ReadKey();
        }
        public static FileInfo GetChromeExe()
        {
            string ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string ProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            string chromePath = @"Google\Chrome\Application\chrome.exe";
            string chromePathx64 = Path.Combine(ProgramFiles, chromePath);
            string chromePathx86 = Path.Combine(ProgramFilesX86, chromePath);
            FileInfo chromeExe = null;
            if (File.Exists(chromePathx64))
                chromeExe = new FileInfo(chromePathx64);
            else if (File.Exists(chromePathx86))
                chromeExe = new FileInfo(chromePathx86);
            return chromeExe;
        }
        public static IEnumerable<string> GetChromeShortcutPaths()
        {
            string ApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            List<string> chromeShortCutpaths = new List<string>
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), @"Google Chrome.lnk"), // common start menu
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), @"Google Chrome.lnk"), // start menu
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"Google Chrome.lnk"),
                Path.Combine(ApplicationData, @"Microsoft\Internet Explorer\Quick Launch\Google Chrome.lnk"),
            };

            string taskbarPath = Path.Combine(ApplicationData, @"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar");
            if (Directory.Exists(taskbarPath))
            {
                chromeShortCutpaths.AddRange(
                Directory.GetFiles(taskbarPath, "*chrome*", SearchOption.AllDirectories));
            }

            return chromeShortCutpaths;
        }
        public static IWshShortcut CreateNewShortcut(string shortcutPath, string targetPath, string iconPath)
        {
            IWshRuntimeLibrary.WshShell shell = new WshShell();
            IWshShortcut shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = targetPath;
            shortcut.IconLocation = iconPath;
            return shortcut;
        }
        public static void SaveShortcustWithArguments(string shortcutPath, FileInfo chromeExe)
        {
            IWshShortcut shortcut = CreateNewShortcut(shortcutPath, chromeExe.FullName, chromeExe.FullName);
            var appUrl = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Config["App"]));
            shortcut.Arguments = $"--start-fullscreen --profile-directory=\"Default\" --app=\"{appUrl}\"";
            shortcut.Save();
        }
        public static void SaveShortcustWithoutArguments(string shortcutPath, string targetPath, string iconPath)
        {
            IWshRuntimeLibrary.WshShell shell = new WshShell();
            IWshShortcut shortcut = CreateNewShortcut(shortcutPath, targetPath, iconPath);
            shortcut.Arguments = "";
            shortcut.Save();
        }
        public static bool IsConnectedToInternet()
        {
            try
            {
                string url = "https://www.google.com";
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.KeepAlive = false;
                request.Timeout = 3000;
                var response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                    Task.Delay(5000);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
