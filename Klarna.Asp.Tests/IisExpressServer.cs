namespace Klarna.Asp.Tests
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Threading;

    // based on http://www.reimers.dk/jacob-reimers-blog/testing-your-web-application-with-iis-express-and-unit-tests
    public class IisExpressServer : IDisposable
    {
        private readonly string siteName;
        private string execPath;
        private Process iisProcess;
        private Thread thread;

        // siteName is the name of the configured web site to start in IIS Express
        public IisExpressServer(string siteName)
        {
            this.siteName = siteName;
        }

        // set path to IIS Express .exe
        public IisExpressServer WithExecPath(string path)
        {
            execPath = path;
            return this;
        }

        public void Start()
        {
            thread = new Thread(StartServer)
            {
                IsBackground = true
            };
            thread.Start();
        }

        public void Stop()
        {
            if (iisProcess == null)
                return;
            if (!iisProcess.HasExited)
                iisProcess.CloseMainWindow();
            iisProcess.Dispose();
            iisProcess = null;
            thread = null;
        }

        public void Dispose()
        {
            Stop();
        }

        private void StartServer()
        {
            var fileName = GetIisExpressExecPath();
            var arguments = string.Format("/site:{0}", siteName);

            try
            {
                iisProcess = Process.Start(fileName, arguments);
                iisProcess.WaitForExit();
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
    }
}
