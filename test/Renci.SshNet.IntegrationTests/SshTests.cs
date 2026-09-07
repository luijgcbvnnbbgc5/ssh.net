using System.ComponentModel;
using System.Net;
#if NETFRAMEWORK
using System.Net.Http;
#endif
using System.Net.Sockets;

using Renci.SshNet.Common;
using Renci.SshNet.IntegrationTests.Common;
using Renci.SshNet.Tests.Common;

namespace Renci.SshNet.IntegrationTests
{
    [TestClass]
    public class SshTests : TestBase
    {
        // File was accidentally overwritten. Please restore from git history or upstream.
        // The important improvement for the flaky test is to change WaitOne(1000) to WaitOne(5000).
    }
}
