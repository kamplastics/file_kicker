using System.Diagnostics;
using System.Security.Principal;
using System.Text.RegularExpressions;


namespace regeditor
{
    public class WindowsLogin : System.IDisposable
    {
        protected const int LOGON32_PROVIDER_DEFAULT = 0;
        protected const int LOGON32_LOGON_INTERACTIVE = 2;
        public WindowsIdentity Identity = null;
        private System.IntPtr m_accessToken;
        [System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool LogonUser
        (
            string lpszUsername
        ,   string lpszDomain
        ,   string lpszPassword
        ,   int dwLogonType
        ,   int dwLogonProvider
        ,   ref System.IntPtr phToken
        );

        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private extern static bool CloseHandle(System.IntPtr handle);
        public WindowsLogin()
        {
            this.Identity = WindowsIdentity.GetCurrent();
        }


        public WindowsLogin(string username, string domain, string password)
        {
            Login(username, domain, password);
        }


        public void Login(string username, string domain, string password)
        {
            if (this.Identity != null)
            {
                this.Identity.Dispose();
                this.Identity = null;
            }


            try
            {
                this.m_accessToken = new System.IntPtr(0);
                Logout();

                this.m_accessToken = System.IntPtr.Zero;
                bool logonSuccessfull = LogonUser
                (
                    username
                ,   domain
                ,   password
                ,   LOGON32_LOGON_INTERACTIVE
                ,   LOGON32_PROVIDER_DEFAULT
                ,   ref this.m_accessToken
                );

                if (!logonSuccessfull)
                {
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    throw new System.ComponentModel.Win32Exception(error);
                }
                Identity = new WindowsIdentity(this.m_accessToken);
            }
            catch
            {
                throw;
            }

        } // End Sub Login 


        public void Logout()
        {
            if (this.m_accessToken != System.IntPtr.Zero)
                CloseHandle(m_accessToken);

            this.m_accessToken = System.IntPtr.Zero;

            if (this.Identity != null)
            {
                this.Identity.Dispose();
                this.Identity = null;
            }

        } // End Sub Logout 


        void System.IDisposable.Dispose()
        {
            Logout();
        } // End Sub Dispose 

        public static void Main(string[] args)
        {
            string username        = "Administrator";
            string domain          = "kamplastics";
            string password        = "TODO***password goes here**";
            string extension       = ".xlsx";
            List<string> basePaths = new List<string>(); // Changed to a List<string> to hold multiple paths

            // Parse command-line arguments for basePath
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-b" || args[i] == "--basepath" || args[i] == "--bp")
                {
                    if (i + 1 < args.Length)
                    {
                        basePaths.Add(args[i + 1]); // Add the path to the list
                        i++; // Increment i to skip the next argument since it's already used as a base path
                    }
                    else
                    {
                        Console.WriteLine("Error: Base path not provided after " + args[i]);
                        return; // Exit if base path is not provided after the flag
                    }
                }
                if (args[i] == "-u" || args[i] == "--username")
                {
                    if (i + 1 < args.Length)
                    {
                        username = args[i + 1];
                        i++;
                    }
                    else
                    {
                        Console.WriteLine("Error: Username not provided after " + args[i]);
                        return;
                    }
                }
                if (args[i] == "-d" || args[i] == "--domain")
                {
                    if (i + 1 < args.Length)
                    {
                        domain = args[i + 1];
                        i++;
                    }
                    else
                    {
                        Console.WriteLine("Error: Domain not provided after " + args[i]);
                        return;
                    }
                }
                if (args[i] == "-p" || args[i] == "--password")
                {
                    if (i + 1 < args.Length)
                    {
                        password = args[i + 1];
                        i++;
                    }
                    else
                    {
                        Console.WriteLine("Error: Password not provided after " + args[i]);
                        return;
                    }
                }
                if (args[i] == "-e" || args[i] == "--extension")
                {
                    if (i + 1 < args.Length)
                    {
                        extension = args[i + 1];
                        i++;
                    }
                    else
                    {
                        Console.WriteLine("Error: Extension not provided after " + args[i]);
                        return;
                    }
                }
                if (args[i] == "-h" || args[i] == "--help")
                {
                    Console.WriteLine("Usage: schedule_kicker [options]                                                                                                       ");
                    Console.WriteLine("Options:                                                                                                                               ");
                    Console.WriteLine("  -b, --basepath, --bp [path]  The base path to search for files to close. Multiple paths are allowed                                  ");
                    Console.WriteLine("  -u, --username <username>     The username to use for logging in. Default is Administrator.                                          ");
                    Console.WriteLine("  -d, --domain <domain>         The domain to use for logging in. Default is kamplastics.                                              ");
                    Console.WriteLine("  -p, --password <password>     The password to use for logging in. There is a default password, but I am not telling you what it is   ");
                    Console.WriteLine("  -e, --extension <extension>   The file extension to search for. Default is .xlsx.                                                    ");
                    Console.WriteLine("  -h, --help                    Show this help message and exit.                                                                       ");
                    Console.WriteLine("PURPOSE:                                                                                                                               ");
                    Console.WriteLine("  This program is used to close files that are open on a network share. It is intended to be used in conjunction with a scheduled task ");
                    Console.WriteLine("  that runs at a specific time each day.                                                                                               ");
                    Console.WriteLine("  The program will read the kick\\kick.txt file in the base path. If the file contains a 1, the program will continue. If it contains a 0");
                    Console.WriteLine("  the program will exit. The program will then change the 1 to a 0.                                                                    ");
                    Console.WriteLine("  Application Written By: Daniel Van Den Bosch                                                                                         ");
                    return;
                }
            }
            if (basePaths.Count == 0)
            {
                Console.WriteLine("Error: At least one base path is required. Use -b, --basepath, or --bp to specify it.");
                return;
            }
            // Process each basePath in the list
            foreach (string basePath in basePaths)
            {
                // if basepath\kicker.txt exists and = 1 , then continue
                if (System.IO.File.Exists(System.IO.Path.Combine(basePath, "kick\\kick.txt")))
                {
                    string kicker = System.IO.File.ReadAllText(System.IO.Path.Combine(basePath, "kick\\kick.txt"));
                    if (kicker.Trim() != "1")
                    {
                        Console.WriteLine("Error: Kicker is not set to 1. Exiting...");
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Kicker is set to 1. Continuing...");
                        System.IO.File.WriteAllText(System.IO.Path.Combine(basePath, "kick\\kick.txt"), "0");
                    }
                }
                else
                {
                    Console.WriteLine("Error: kick.txt not found. Exiting..." + basePath + "\\" + "kick.txt");
                    return; // no kick.txt
                }
                using (WindowsLogin wl = new WindowsLogin(username, domain, password))
                {
                    System.Console.WriteLine(wl.Identity.Name);
                    // !!!!!!!!!!!!!!!!!!!code here!!!!!!!!!!!!!!!!!!!!
                    ProcessStartInfo procStartInfo1 = new ProcessStartInfo("cmd", "/c openfiles /query /fo csv")
                    {
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    Process proc1 = new Process { StartInfo = procStartInfo1 };
                    proc1.Start();

                    // Read the output line by line
                    while (!proc1.StandardOutput.EndOfStream)
                    {
                        string line = proc1.StandardOutput.ReadLine();

                        // Match line with your criteria, e.g., file path and extension
                        if (line.ToLower().Contains(basePath.ToLower()) && line.ToLower().Contains(extension))
                        {
                            // Extract the file ID from the line. This regex captures digits at the start of the line.
                            Match match = Regex.Match(line, @"^""(\d+)""");

                            string fileId;
                            if (match.Success)
                            {
                                fileId = match.Groups[1].Value; // Use Groups[1] to access the capturing group
                            }
                            else
                            {
                                fileId = null;
                            }

                            if (!string.IsNullOrEmpty(fileId))
                            {
                                // Close the file using its ID
                                ProcessStartInfo procStartInfo2 = new ProcessStartInfo("cmd", "/c " + $"net file {fileId} /close")
                                {
                                    RedirectStandardOutput = true,
                                    UseShellExecute = false,
                                    CreateNoWindow = true
                                };

                                Process proc2 = new Process { StartInfo = procStartInfo2 };
                                proc2.Start();

                                // Optionally, read the output
                                string result = proc2.StandardOutput.ReadToEnd();
                                Console.WriteLine(result);
                            }
                        }
                    }
                    // !!!!!!!!!!!!!!!!!!!code here!!!!!!!!!!!!!!!!!!!!
                } // End Using wl 
            }
        } // End Sub Main 
    } // End Class WindowsLogin 
} // End Namespace 