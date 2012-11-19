using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace AspUnitRunner.Sample.Tests.NUnit
{
    // based on http://www.reimers.dk/jacob-reimers-blog/testing-your-web-application-with-iis-express-and-unit-tests
    public class IisExpressServer : IDisposable
    {
        private readonly string _siteName;
        private string _execPath;
        private Process _iisProcess;
        private Thread _thread;

        // siteName is the name of the configured web site to start in IIS Express
        public IisExpressServer(string siteName)
        {
            _siteName = siteName;
        }

        // set path to IIS Express .exe
        public IisExpressServer WithExecPath(string path)
        {
            _execPath = path;
            return this;
        }

        public void Start()
        {
            _thread = new Thread(StartServer)
            {
                IsBackground = true
            };
            _thread.Start();
        }

        public void Stop()
        {
            if (_iisProcess == null)
                return;
            if (!_iisProcess.HasExited)
                _iisProcess.CloseMainWindow();
            _iisProcess.Dispose();
            _iisProcess = null;
            _thread = null;
        }

        public void Dispose()
        {
            Stop();
        }

        private void StartServer()
        {
            var fileName = GetIisExpressExecPath();
            var arguments = string.Format("/site:{0}", _siteName);

            try
            {
                _iisProcess = Process.Start(fileName, arguments);
                _iisProcess.WaitForExit();
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
            if (string.IsNullOrEmpty(_execPath))
                return Path.Combine(GetProgramFilesDir(), @"IIS Express\iisexpress.exe");
            return _execPath;
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
