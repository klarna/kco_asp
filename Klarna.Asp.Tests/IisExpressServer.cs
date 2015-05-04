namespace Klarna.Asp.Tests
{
    using NUnit.Framework;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Threading;

    // based on http://www.reimers.dk/jacob-reimers-blog/testing-your-web-application-with-iis-express-and-unit-tests
    [SetUpFixture]
    public class IisExpressServer : IDisposable
    {
        private readonly string path;
        private readonly int port;
        private string execPath;
        private Process iisProcess;
        private Thread thread;

        public IisExpressServer()
        {
            string cwd = Directory.GetCurrentDirectory();
            DirectoryInfo dir = Directory.GetParent(cwd).Parent.Parent;
            this.path = dir.FullName + "\\Klarna.Asp.Test.Web";
            this.port = 54979;
        }

        // siteName is the name of the configured web site to start in IIS Express
        public IisExpressServer(string path)
        {
            this.path = path;
            this.port = 54979;
        }

        // set path to IIS Express .exe
        public IisExpressServer WithExecPath(string path)
        {
            execPath = path;
            return this;
        }

        [SetUp]
        public void Start()
        {
            thread = new Thread(StartServer)
            {
                IsBackground = true
            };
            thread.Start();
        }

        [TearDown]
        public void Stop()
        {
            if (iisProcess == null)
                return;
            if (!iisProcess.HasExited)
                iisProcess.CloseMainWindow();
            iisProcess.Kill();
            iisProcess.Close();
            iisProcess.Dispose();
            iisProcess = null;
            thread.Abort();
            thread = null;
        }

        public void Dispose()
        {
            Stop();
        }

        private void StartServer()
        {
            var fileName = GetIisExpressExecPath();
            var arguments = string.Format("/path:\"{0}\" /port:{1} /systray:false", path, port);

            Debug.Write(fileName + " ");
            Debug.WriteLine(arguments);

            try
            {
                iisProcess = new Process();
                iisProcess.StartInfo.UseShellExecute = false;
                iisProcess.StartInfo.RedirectStandardOutput = true;
                iisProcess.StartInfo.RedirectStandardError = true;
                iisProcess.StartInfo.FileName = fileName;
                iisProcess.StartInfo.Arguments = arguments;

                iisProcess.OutputDataReceived += new DataReceivedEventHandler(OutputHandler);
                iisProcess.ErrorDataReceived += new DataReceivedEventHandler(OutputHandler);

                iisProcess.Start();

                iisProcess.BeginOutputReadLine();
                iisProcess.BeginErrorReadLine();
            }
            catch
            {
                Stop();
                throw;
            }
        }

        // assumes IIS Express is installed at %programfiles(x86)%\IIS Express\iisexpress.exe or
        // %programfiles%\IIS Express\iisexpress.exe
        private string GetIisExpressExecPath()
        {
            if (string.IsNullOrEmpty(execPath))
                return Path.Combine(GetProgramFilesDir(), @"IIS Express\iisexpress.exe");
            return execPath;
        }

        private static string GetProgramFilesDir()
        {
            // note: in .NET 4, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) can be used instead
            var programFiles = Environment.GetEnvironmentVariable("programfiles(x86)");
            if (string.IsNullOrEmpty(programFiles))
                return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            return programFiles;
        }

        private static void OutputHandler(
            Object sender,
            DataReceivedEventArgs e
        ) {
            if (!String.IsNullOrEmpty(e.Data))
            {
                Debug.WriteLine(e.Data);
            }
        }
    }
}
