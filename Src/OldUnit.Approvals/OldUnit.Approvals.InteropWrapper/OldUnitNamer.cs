namespace OldUnit.Approvals.InteropWrapper
{
	public class OldUnitNamer : ApprovalTests.Core.IApprovalNamer
	{
	    public OldUnitNamer(string name, string sourcePath)
	    {
	        Name = name;
	        SourcePath = sourcePath;
	    }

	    public string Name{get; private set; }

        public string SourcePath {get; private set; }

    }


}
