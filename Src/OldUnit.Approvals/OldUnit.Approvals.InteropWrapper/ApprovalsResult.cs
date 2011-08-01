using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace OldUnit.Approvals.InteropWrapper
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DD5FCC50-BDED-409D-8E88-FCC51D02A118")]
    public class ApprovalsResult : IApprovalsResult
    {
        public bool Approved { get; private set; }
        public string ReceivedFile { get; private set; }
        public string ApprovedFile { get; private set; }
        public string FixtureName { get; private set; }
        public string TestMethodName { get; private set; }
        public string ApproveFileCommandText
        {
            get { return string.Format("move /Y \"{0}\" \"{1}\"", ReceivedFile, ApprovedFile); }
        }

        public ApprovalsResult(bool approved, string receivedFile, string approvedFile, string fixtureName, string testMethodName)
        {
            Approved = approved;
            ReceivedFile = receivedFile;
            ApprovedFile = approvedFile;
            FixtureName = fixtureName;
            TestMethodName = testMethodName;
        }


        

        
    }
}
