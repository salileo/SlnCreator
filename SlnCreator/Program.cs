using EnvDTE;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace SlnCreator
{
    class Program
    {
        private static int MaxRetryCount = 5;
        private static TimeSpan GapBetweenRetries = TimeSpan.FromSeconds(5);

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Log("Invalid arguments");
                return;
            }

            string projectFileList = args[0];
            string solutionFilePath = args[1];
            string solutionFileName = args[2];

            try
            {
                CreateSolutionFileV2(projectFileList, solutionFilePath, solutionFileName);
            }
            catch (AggregateException ex)
            {
                LogException(ex);
            }
        }

        [STAThread]
        private static void CreateSolutionFile(string projectFileList, string solutionFilePath, string solutionFileName)
        {
            Log("Starting VS...");
            DTE vsObj = StartVisualStudio();

            try
            {
                bool opened = false;
                bool updated = false;
                string solutionFile = Path.Combine(solutionFilePath, solutionFileName + ".sln");

                Solution slnObj = GetSolutionInterface(vsObj);

                try
                {
                    if (File.Exists(solutionFile))
                    {
                        Log(string.Format("Opening solution file {0} ...", solutionFile));
                        OpenSolutionFile(slnObj, solutionFile);
                    }
                    else
                    {
                        Log(string.Format("Create solution file {0} under {1}...", solutionFileName, solutionFilePath));
                        CreateSolutionFile(slnObj, solutionFilePath, solutionFileName);
                        updated = true;
                    }

                    opened = true;
                    if (!File.Exists(projectFileList))
                    {
                        return;
                    }

                    using (FileStream stream = File.OpenRead(projectFileList))
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            while (!reader.EndOfStream)
                            {
                                string projectFile = reader.ReadLine();

                                if (projectFile != null)
                                {
                                    projectFile = projectFile.Trim();
                                }

                                if (!string.IsNullOrEmpty(projectFile))
                                {
                                    try
                                    {
                                        updated |= AddProject(slnObj, projectFile);
                                    }
                                    catch (AggregateException ex)
                                    {
                                        LogException(ex);
                                        throw;
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    if (updated)
                    {
                        try
                        {
                            Log("Saving solution file...");
                            SaveSolutionFile(slnObj, solutionFile);
                        }
                        catch (AggregateException ex)
                        {
                            LogException(ex);
                        }
                    }

                    if (opened)
                    {
                        try
                        {
                            Log("Closing solution file...");
                            CloseSolutionFile(slnObj);
                        }
                        catch (AggregateException ex)
                        {
                            LogException(ex);
                        }
                    }
                }
            }
            finally
            {
                try
                {
                    Log("Stopping VS...");
                    StopVisualStudio(vsObj);
                }
                catch (AggregateException ex)
                {
                    LogException(ex);
                }
            }
        }

        [STAThread]
        private static void CreateSolutionFileV2(string projectFileList, string solutionFilePath, string solutionFileName)
        {
            string solutionFile = Path.Combine(solutionFilePath, solutionFileName + ".sln");

            if (!File.Exists(solutionFile))
            {
                Log("Starting VS...");
                DTE vsObj = StartVisualStudio();

                bool failure = false;
                try
                {
                    Solution slnObj = GetSolutionInterface(vsObj);
                    bool opened = false;

                    try
                    {
                        Log(string.Format("Create solution file {0} under {1}...", solutionFileName, solutionFilePath));
                        CreateSolutionFile(slnObj, solutionFilePath, solutionFileName);
                        opened = true;
                    }
                    catch (AggregateException ex)
                    {
                        LogException(ex);
                        failure = true;
                    }
                    finally
                    {
                        if (opened)
                        {
                            try
                            {
                                Log("Saving solution file...");
                                SaveSolutionFile(slnObj, solutionFile);
                            }
                            catch (AggregateException ex)
                            {
                                LogException(ex);
                                failure = true;
                            }

                            try
                            {
                                Log("Closing solution file...");
                                CloseSolutionFile(slnObj);
                            }
                            catch (AggregateException ex)
                            {
                                LogException(ex);
                                failure = true;
                            }
                        }
                    }
                }
                finally
                {
                    try
                    {
                        Log("Stopping VS...");
                        StopVisualStudio(vsObj);
                    }
                    catch (AggregateException ex)
                    {
                        LogException(ex);
                        failure = true;
                    }
                }

                if (failure)
                {
                    return;
                }
            }

            if (!File.Exists(projectFileList))
            {
                return;
            }

            using (FileStream stream = File.OpenRead(projectFileList))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    while (!reader.EndOfStream)
                    {
                        string projectFile = reader.ReadLine();

                        if (projectFile != null)
                        {
                            projectFile = projectFile.Trim();
                        }

                        if (!string.IsNullOrEmpty(projectFile))
                        {
                            try
                            {
                                AddProjectV2(solutionFile, projectFile);
                            }
                            catch (AggregateException ex)
                            {
                                LogException(ex);
                                throw;
                            }
                        }
                    }
                }
            }
        }

        [STAThread]
        private static DTE StartVisualStudio()
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();
            DTE vsObj = null;

            string vsString = "VisualStudio.DTE.12.0";
            if (File.Exists(@"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe"))
            {
                vsString = "VisualStudio.DTE.14.0";
            }

            while (vsObj == null && retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                   System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    object obj = Activator.CreateInstance(Type.GetTypeFromProgID(vsString, true), true);
                    vsObj = (DTE)obj;
                    if (vsObj == null)
                    {
                        retryCount++;
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            if (vsObj == null)
            {
                throw new AggregateException("Giving up trying to start VS !!!", exceptions);
            }

            return vsObj;
        }

        [STAThread]
        private static void StopVisualStudio(DTE vsObj)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            while (vsObj != null && retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    vsObj.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, vsObj, null);
                    vsObj = null;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            if (vsObj != null)
            {
                throw new AggregateException("Giving up trying to close VS !!!", exceptions);
            }
        }

        [STAThread]
        private static Solution GetSolutionInterface(DTE vsObj)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();
            Solution slnObj = null;

            while (slnObj == null && retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj = vsObj.Solution;
                    if (slnObj == null)
                    {
                        retryCount++;
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            if (slnObj == null)
            {
                throw new AggregateException("Giving up trying to get solution interface !!!", exceptions);
            }

            return slnObj;
        }

        [STAThread]
        private static void CreateSolutionFile(Solution slnObj, string path, string filename)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj.Create(path, filename);
                    return;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            throw new AggregateException("Giving up trying to create solution !!!", exceptions);
        }

        [STAThread]
        private static void OpenSolutionFile(Solution slnObj, string path)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj.Open(path);
                    return;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            throw new AggregateException("Giving up trying to open solution !!!", exceptions);
        }

        [STAThread]
        private static void SaveSolutionFile(Solution slnObj, string path)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj.SaveAs(path);
                    return;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            throw new AggregateException("Giving up trying to save solution !!!", exceptions);
        }

        [STAThread]
        private static void CloseSolutionFile(Solution slnObj)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj.Close(false);
                    return;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            throw new AggregateException("Giving up trying to close solution !!!", exceptions);
        }

        [STAThread]
        private static bool AddProject(Solution slnObj, string path)
        {
            int retryCount = 0;
            List<Exception> exceptions = new List<Exception>();

            if (!File.Exists(path))
            {
                Log(string.Format("Could not find project file {0} !!!", path));
                return false;
            }

            retryCount = 0;
            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    foreach (Project project in slnObj.Projects)
                    {
                        if (string.Equals(project.FullName, path, StringComparison.InvariantCultureIgnoreCase))
                        {
                            Log(string.Format("Already added project file {0} !!!", path));
                            return false;
                        }
                    }

                    break;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            retryCount = 0;
            while (retryCount < MaxRetryCount)
            {
                if (retryCount > 0)
                {
                    System.Threading.Thread.Sleep(GapBetweenRetries);
                }

                try
                {
                    slnObj.AddFromFile(path, false);
                    Log(string.Format("Added project file {0} !!!", path));
                    return true;
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    retryCount++;
                }
            }

            throw new AggregateException(string.Format("Giving up trying to add project file {0} !!!", path), exceptions);
        }

        [STAThread]
        private static bool AddProjectV2(string solutionFile, string projectFile)
        {
            List<Exception> exceptions = new List<Exception>();

            if (!File.Exists(projectFile))
            {
                Log(string.Format("Could not find project file {0} !!!", projectFile));
                return false;
            }

            string solutionFileData = File.ReadAllText(solutionFile);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(projectFile);

            XmlNode guidNode = null;
            XmlNodeList list = xmlDoc.GetElementsByTagName("ProjectGuid");
            if (list.Count == 1)
            {
                guidNode = list[0].FirstChild;
            }

            string projectGuid = null;
            if (guidNode != null)
            {
                projectGuid = guidNode.Value;
            }

            if (string.IsNullOrEmpty(projectGuid))
            {
                Log(string.Format("Could not find project guid for file {0} !!!", projectFile));
                return false;
            }

            string projectName = Path.GetFileNameWithoutExtension(projectFile);

            if (solutionFileData.Contains(projectGuid))
            {
                Log(string.Format("Already added project file {0} !!!", projectFile));
                return false;
            }

            string entryData = string.Format(
                "Project(\"{{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}}\") = \"{0}\", \"{1}\", \"{2}\"\r\nEndProject\r\n",
                projectName,
                projectFile,
                projectGuid);

            int index = solutionFileData.IndexOf("Global");
            if (index != -1)
            {
                solutionFileData = solutionFileData.Insert(index, entryData);
            }
            else
            {
                solutionFileData += entryData;
            }

            File.WriteAllText(solutionFile, solutionFileData);
            Log(string.Format("Added project file {0} !!!", projectFile));
            return true;
            //throw new AggregateException(string.Format("Giving up trying to add project file {0} !!!", path), exceptions);
        }

        private static void Log(string data)
        {
            Console.WriteLine(data);
        }

        private static void LogException(Exception ex)
        {
            Log(string.Format("Exception = {0}", ex));
        }
    }
}
