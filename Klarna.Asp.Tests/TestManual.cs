using System.Text;
using NUnit.Framework;
using AspUnitRunner;

namespace AspUnitRunner.Sample.Tests.NUnit
{
    // demonstrates manually specifying test container names for an ASPUnit test suite
    // (one NUnit test per ASPUnit container)
    [TestFixture]
    public class TestManual
    {
        // set the URL for your ASPUnit tests
        private const string AspTestUrl = "http://localhost:54979/Tests/Default.asp";
        // set the site name as configured in IIS Express (defaults to name of sample web project: AspUnitRunner.Sample.Web)
        private const string AspSiteName = "Klarna.Asp.Tests.Web";

        private IisExpressServer _iisServer;

        [Test]
        public void TestContainer(
            // set ASPUnit test containers here
            [Values("StringUtilityTest", "FailureTest")] string testContainer
        )
        {
            var runner = Runner.Create(AspTestUrl)
                .WithEncoding(Encoding.UTF8).WithTestContainer(testContainer);
            var results = runner.Run();

            // this results in slightly cleaner output than Assert.That(results.Successful...)
            if (!results.Successful)
                Assert.Fail(results.Format());

            if (results.Tests == 0)
                Assert.Inconclusive("0 tests were run");
        }

        [TestFixtureSetUp]
        public void StartServer()
        {
            //_iisServer = new IisExpressServer(AspSiteName);
            //_iisServer.Start();
        }

        [TestFixtureTearDown]
        public void StopServer()
        {
            //_iisServer.Stop();
        }
    }
}
