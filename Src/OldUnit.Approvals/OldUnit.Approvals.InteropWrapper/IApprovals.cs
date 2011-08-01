using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace OldUnit.Approvals.InteropWrapper
{
    [Guid("A7BDF162-8C75-4024-8D86-1E8BA804E2A4")]
    public interface IApprovals
    {
        IApprovalsResult ApproveString(string stringToApprove);

        void UseDiffReporter();

        void UseSilentReporter();
    }
}
