using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OldUnit.Approvals.InteropWrapper
{
    public interface IFilePathFinder
    {
        string Find(string fileName);
    }
}
