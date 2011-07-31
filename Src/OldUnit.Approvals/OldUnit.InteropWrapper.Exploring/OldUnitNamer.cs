
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OldUnit.InteropWrapper.Exploring
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
