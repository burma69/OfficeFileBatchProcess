using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Linq;


namespace OfficeBatchProcess
{
    class ProcessFileThreadState
    {
        public string FromFile, ToFile;        

        public ProcessFileThreadState(string FromFileName, string ToFileName)
        {
            FromFile = FromFileName;
            ToFile = ToFileName;
        }
    }

    class Program
    {        
        static void ProcessFile(Object state)
        {
            try
            {
                WorkThreadLimit.WaitOne(); 

                using (FileProcessor x = new FileProcessor())
                {
                    ProcessFileThreadState s = (state as ProcessFileThreadState);

                    if (s.FromFile.EndsWith(".pdf"))
                    {
                        /*
                        byte[] DOCXBuffer = Convert(s.FromFile);
                        string DOCXSource = s.FromFile.Substring(0, s.FromFile.LastIndexOf(".")) + ".docx";

                        Interlocked.Add(ref FilesCount, 1);

                        Interlocked.Add(ref ChangesCount, x.ProcessFile(DOCXBuffer, DOCXSource, s.ToFile));
                        */
                    } else
                    {
                        Interlocked.Add(ref FilesCount, 1);

                        Interlocked.Add(ref ChangesCount, x.ProcessFile(s.FromFile, s.ToFile));
                    }
                }
            }
            finally
            {
                WorkThreadLimit.Release();
                finished.Signal();
            }
        }

        static int ChangesCount = 0;
        static int FilesCount = 0;
        static CountdownEvent finished = new CountdownEvent(1);
        static System.Threading.Semaphore WorkThreadLimit;
        static string SourceDirectory;
        static string DestDirectory;


        static void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            // First, process all the files directly under this folder 
            try
            {
                files = (System.IO.FileInfo[])root.GetFiles("*.*").Where(s => s.Name.EndsWith(".xlsx")).ToArray();
            }
            // This is thrown if even one of the files requires permissions greater 
            // than the application provides. 
            catch (UnauthorizedAccessException e)
            {
                // This code just writes out the message and continues to recurse. 
                // You may decide to do something different here. For example, you 
                // can try to elevate your privileges and access the file again.
                Console.WriteLine(e.Message);
            }

            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }

            if (files != null)
            {
                foreach (System.IO.FileInfo fi in files)
                {
                    string DestFileName = DestDirectory + fi.FullName.Substring(fi.FullName.IndexOf(SourceDirectory) + SourceDirectory.Length);

                    DestFileName = DestFileName.Substring(0, DestFileName.LastIndexOf(".")) + ".xlsx";

                    if (File.Exists(DestFileName))
                    {
                        Console.WriteLine("File {0} already exists, skipping", DestFileName);
                        continue;

                    }

                    string ExtractsFileName = DestFileName + " - extracts.xml";

                    if (File.Exists(ExtractsFileName))
                    {
                        Console.WriteLine("File {0} already exists, skipping", ExtractsFileName);
                        continue;

                    }

                    finished.AddCount();
                    System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(ProcessFile), new ProcessFileThreadState(fi.FullName, DestFileName));
                    //ChangesCount += x.ProcessFile(fi.FullName, );
                }

                // Now find all the subdirectories under this directory.
                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    // Resursive call for each subdirectory.
                    WalkDirectoryTree(dirInfo);
                }
            }
        }

        [STAThread]
        static void Main(string[] args)
        {   
            if (args.GetLength(0) == 2)
            using (FileProcessor x = new FileProcessor())
            {
                ChangesCount = x.ProcessFile(args[0], args[1]);
                return;
            }
          

            Console.WriteLine("Automatic Office files update. Syntax: " + Assembly.GetExecutingAssembly().GetName().Name  + " <source> <destination>");
            Console.WriteLine("If <source> and <destination> are not specified, the program will convert all *.xlsx files from SourceDirectory to DestinationDirectory as specified in app.config file");

            Console.WriteLine("Processing with a maximum of " +ConfigurationManager.AppSettings["MaxThreads"] + " threads.");
            System.Threading.ThreadPool.SetMaxThreads(int.Parse(ConfigurationManager.AppSettings["MaxThreads"]), 0);
            WorkThreadLimit = new System.Threading.Semaphore(int.Parse(ConfigurationManager.AppSettings["MaxThreads"]), int.Parse(ConfigurationManager.AppSettings["MaxThreads"]));


            CommonOpenFileDialog dialog = new CommonOpenFileDialog();

            if (ConfigurationManager.AppSettings["SourceDirectory"] == null)
            {
                dialog.Title = "Select a folder with the source XLSX files";
                dialog.IsFolderPicker = true;
                dialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok) SourceDirectory = dialog.FileName; else return;
            }

            if (ConfigurationManager.AppSettings["DestinationDirectory"] == null)
            {
                dialog.Title = "Select the destination folder to store processed files";
                dialog.IsFolderPicker = true;                
                dialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok) DestDirectory = dialog.FileName; else return;
            }

            DirectoryInfo rootDir = new DirectoryInfo(SourceDirectory);
            WalkDirectoryTree(rootDir);

            finished.Signal();
            finished.Wait();            
            Console.WriteLine(ChangesCount.ToString() + " change(s) made total.");
            Console.WriteLine(FilesCount.ToString() + " files processed.");
            Console.WriteLine("[{0}] Finished", DateTime.Now);
            Console.ReadLine();


        }
    }
}
