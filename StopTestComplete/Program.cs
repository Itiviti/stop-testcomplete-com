using System.Runtime.InteropServices;
using System;
using System.Threading;

namespace StopTestComplete
{
    class Program
    {
        public static int Main(string[] args)
        {
            string[] progIDs = new string[] {
                "TestComplete.TestCompleteApplication",
                "TestComplete.TestCompleteApplication.12",
                "TestComplete.TestCompleteX64Application",
                "TestComplete.TestCompleteX64Application.12",
                "TestComplete.TestExecuteApplication",
                "TestComplete.TestExecuteX64Application",
                "TestExecute.TestExecuteApplication",
                "TestExecute.TestExecuteApplication.12",
                "TestExecute.TestExecuteX64Application",
                "TestExecute.TestExecuteX64Application.12"
            };

            string tcProgID = null;
            object testCompleteObject = null;

            foreach(string progId in progIDs)
            {
                try
                {
                    testCompleteObject = Marshal.GetActiveObject(progId);
                    tcProgID = progId;
                    break;
                }
                catch
                {
                    //do nothing and try the next one
                }
            }

            if (testCompleteObject == null)
            {
                Console.WriteLine("Failed to connect to TestComplete/TestExecute.");
                return 2;
            }
            Console.WriteLine("Successfully connected to " + tcProgID);

            // Obtains ITestCompleteCOMManager
            TestComplete.ITestCompleteCOMManager testCompleteManager = (TestComplete.ITestCompleteCOMManager)testCompleteObject;
            // Obtains Integration object
            TestComplete.ItcIntegration integrationObject = testCompleteManager.Integration;

            if (integrationObject.IsRunning())
            {
                Console.WriteLine("TestComplete is executing tests. Stopping it ...");
                stopTcAndWait(integrationObject);
                Console.WriteLine("TestComplete was stopped.");
            }
            else
            {
                Console.WriteLine("TestComplete is not running any tests.");
                return 1;
            }
            Marshal.ReleaseComObject(testCompleteObject);
            return 0;
        }

        private static void stopTcAndWait(TestComplete.ItcIntegration integrationObject)
        {
            integrationObject.Stop();
            try
            {
                while (integrationObject.IsRunning())
                {
                    Console.WriteLine("Waiting for TC to stop ...");
                    Thread.Sleep(2000);
                }
            }
            catch { }
        }
    }
}
