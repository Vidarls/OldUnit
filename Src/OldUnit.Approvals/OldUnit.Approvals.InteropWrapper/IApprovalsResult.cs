using System.Runtime.InteropServices;

namespace OldUnit.Approvals.InteropWrapper
{
    [Guid("571E4A4F-C8C7-42AD-B101-ABF664384FF5")]    
    public interface IApprovalsResult
    {
        bool Approved{ get; }

        string ApprovedFile { get; }

        string ReceivedFile { get; }

        string ApproveFileCommandText { get; }

        string FixtureName { get; }

        string TestMethodName { get; }

    }
}