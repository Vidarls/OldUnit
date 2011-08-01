using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ApprovalTests;

namespace OldUnit.InteropWrapper.Exploring
{
    class Program
    {
        static void Main(string[] args)
        {
            var v = "valuie";
            var x = "xalue;";
            Console.WriteLine("Starting");
            Approvals.DefaultNamerSource = (() => new OldUnitNamer(v, x));
            Approvals.DefaultReporterSource = (() => new ApprovalTests.Reporters.DiffReporter());
            x = @"c:\testnamer";
            Approvals.Approve("TEST");
            
            Console.ReadLine();
        }
    }
}
