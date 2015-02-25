#region Copyright Header
// ----------------------------------------------------------------------------
// <copyright file="OrderTest.cs" company="Klarna AB">
//     Copyright 2013 Klarna AB
//
//     Licensed under the Apache License, Version 2.0 (the "License");
//     you may not use this file except in compliance with the License.
//     You may obtain a copy of the License at
//
//         http://www.apache.org/licenses/LICENSE-2.0
//
//     Unless required by applicable law or agreed to in writing, software
//     distributed under the License is distributed on an "AS IS" BASIS,
//     WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//     See the License for the specific language governing permissions and
//     limitations under the License.
// </copyright>
// <author>Klarna Support: support@klarna.com</author>
// <link>http://developers.klarna.com/</link>
// ----------------------------------------------------------------------------
#endregion
namespace Klarna.Asp.Tests
{
    using System.Text;

    using AspUnitRunner;

    using NUnit.Framework;

    /// <summary>
    /// Tests the ASP Order class.
    /// </summary>
    [TestFixture]
    public class OrderWithConnectorTest
    {
        #region Private Fields

        // set the URL for your ASPUnit tests
        private const string AspTestUrl = "http://localhost:54979/Tests/Default.asp";

        // set the site name as configured in IIS Express
        // (defaults to name of sample web project: AspUnitRunner.Sample.Web)
        private const string AspSiteName = "Klarna.Asp.Tests.Web";

        private IisExpressServer iisServer;

        #endregion

        #region SetUp and TearDown

        [TestFixtureSetUp]
        public void StartServer()
        {
            iisServer = new IisExpressServer(AspSiteName);
            iisServer.Start();
        }

        [TestFixtureTearDown]
        public void StopServer()
        {
            iisServer.Stop();
        }

        #endregion

        /// <summary>
        /// Runs the ASP unit tests for Order class with Connector.
        /// </summary>
        [Test]
        public void OrderWithConnectorTests()
        {
            var runner = Runner.Create(AspTestUrl)
                .WithEncoding(Encoding.UTF8).WithTestContainer("OrderWithConnectorTest");
            var results = runner.Run();

            // this results in slightly cleaner output than Assert.That(results.Successful...)
            if (!results.Successful)
                Assert.Fail(results.Format());

            if (results.Tests == 0)
                Assert.Inconclusive("0 tests were run");
        }
    }
}
