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
    using AspUnitRunner;
    using NUnit.Framework;
    using System.Text;

    /// <summary>
    /// Base test class.
    /// </summary>
    public class TestCase
    {
        // set the URL for your ASPUnit tests
        protected const string AspTestUrl = "http://localhost:54979/Tests/Default.asp";

        /// <summary>
        /// Runs the ASP unit tests in the specified container
        /// </summary>
        /// <param name="container">ASP test container</param>
        protected void RunAspTests(string container)
        {
            var runner = Runner.Create(AspTestUrl)
                .WithEncoding(Encoding.UTF8).WithTestContainer(container);
            var results = runner.Run();

            // this results in slightly cleaner output than Assert.That(results.Successful...)
            if (!results.Successful)
            {
                Assert.Fail(results.Format());
            }

            if (results.Tests == 0)
            {
                Assert.Inconclusive("0 tests were run");
            }
        }

    }
}
