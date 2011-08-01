using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace OldUnit.Approvals.InteropWrapper
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("274EA315-F2C1-44B0-95AF-1B3F76664EAE")]
    public class Approvals : IApprovals
    {
        public IApprovalsResult ApproveString(string stringToApprove)
        {
            throw new NotImplementedException();
        }

        public void UseDiffReporter()
        {
            throw new NotImplementedException();
        }

        public void UseSilentReporter()
        {
            throw new NotImplementedException();
        }
    }
}
