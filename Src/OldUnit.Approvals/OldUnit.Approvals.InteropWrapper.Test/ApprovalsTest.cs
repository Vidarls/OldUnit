using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;

namespace OldUnit.Approvals.InteropWrapper.Test
{
    public class ApprovalsTest
    {
        [Fact]
        [ApprovalTests.Reporters.UseReporter(typeof(InteropWrapper.OldUnitDiffReporter))]
        public void Test_test()
        {
            ApprovalTests.Approvals.Approve("string");
        }
    }
}
