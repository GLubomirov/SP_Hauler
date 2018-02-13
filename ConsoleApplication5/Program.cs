using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ConsoleApplication5
{
    class Program
    {

static void Main(string[] args)
        {
            //nameSpace = the namespace;
            //outDirectory = where the file will be extracted to;
            //internalFilePath = the name of the folder inside visual studio which the files are in;
            //resourceName = the name of the file;
            string AppFolderUser = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string progFolder = AppFolderUser + @"\SP_Copy_Content";
            if (Directory.Exists(progFolder))
            {
                Directory.Delete(progFolder, true);
            }
            System.IO.Directory.CreateDirectory(progFolder);

            //Console.Write(AppFolderUser);
            Extract("ConsoleApplication5", progFolder, "Files", @"Batch.bat");
            Extract("ConsoleApplication5", progFolder, "Files", @"Copy_Home.ps1");
            Extract("ConsoleApplication5", progFolder, "Files", @"Launcher.ps1");
            Extract("ConsoleApplication5", progFolder, "Files", @"Copy_Home.txt");
            Extract("ConsoleApplication5", progFolder, "Files", @"MainWindow.xaml");
            Extract("ConsoleApplication5", progFolder, "Files", @"Microsoft.SharePoint.Client.dll");
            Extract("ConsoleApplication5", progFolder, "Files", @"Microsoft.SharePoint.Client.Runtime.dll");
            Extract("ConsoleApplication5", progFolder, "Files", @"Microsoft.SharePoint.Client.Taxonomy.dll");
            Extract("ConsoleApplication5", progFolder, "Images", @"file-icon-28038.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"icon-folder-128.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Logo.ico");
            Extract("ConsoleApplication5", progFolder, "Images", @"Button_Next_S.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Button_Next_S_Click.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Button_Prev_S.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Button_Prev_S_Click.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Close_Clicked.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Close.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Top_Banner.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Top_Four.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Top_Three.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Top_Two.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Top_One.png");
            Extract("ConsoleApplication5", progFolder, "Images", @"Fields_Back_Advanced.png");
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.FileName = progFolder + "\\"+ "Batch.bat";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.Arguments = "gT4XPfvcJmHkQ5tYjY3fNgi7uwG4FB9j";

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                // Log error.
            }
           //if (Directory.Exists(progFolder))
           //{
           //     Directory.Delete(progFolder, true);
           //}
        }

        public static void Extract(string nameSpace, string outDirectory, string internalFilePath, string resourceName)
        {
            Assembly assembly = Assembly.GetCallingAssembly();

            using (Stream s = assembly.GetManifestResourceStream(nameSpace + "." + (internalFilePath == "" ? "" : internalFilePath + ".") + resourceName))
            using (BinaryReader r = new BinaryReader(s))
            using (FileStream fs = new FileStream(outDirectory + "\\" + resourceName, FileMode.OpenOrCreate))
            using (BinaryWriter w = new BinaryWriter(fs))
            {
                w.Write(r.ReadBytes((int)s.Length));
            }
        }
    }
}

