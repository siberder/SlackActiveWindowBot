using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Slack.NetStandard;
using Slack.NetStandard.WebApi.Users;

namespace SlackActiveWindowBot
{
    class Program
    {
        private static string Token { get; set; }

        private static string OSAScript = "getwindow.script";

        static void Main(string[] args)
        {
            const string tokenFileName = "token.txt";
            try
            {
                Token = File.ReadAllText(tokenFileName);

                if (string.IsNullOrEmpty(Token))
                {
                    throw new FileNotFoundException();
                }
            }
            catch (FileNotFoundException)
            {
                Console.Write("Enter your token: ");
                var token = Console.ReadLine();
                File.WriteAllText(tokenFileName, token);
                Token = token;
            }

            MainLoop();

            Console.ReadLine();
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        private static string GetActiveWindowTitle_OSX()
        {
            string test = $" -c \"osascript {OSAScript} \"";
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                FileName = "/bin/bash",
                Arguments = test,
                CreateNoWindow = false,
                RedirectStandardOutput = true,
                WorkingDirectory = Directory.GetCurrentDirectory()
            };
            process.StartInfo = startInfo;
            process.Start();
            return process?.StandardOutput.ReadToEnd();
        }

        private static string GetActiveWindowTitle_Windows()
        {
            const int nChars = 256;
            var buff = new StringBuilder(nChars);
            var handle = GetForegroundWindow();

            if (GetWindowText(handle, buff, nChars) > 0)
            {
                return buff.ToString();
            }

            return null;
        }

        static async void MainLoop()
        {
            var lastWindowName = string.Empty;

            while (true)
            {
                string windowName;

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    windowName = GetActiveWindowTitle_Windows();
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    windowName = GetActiveWindowTitle_OSX();
                }
                else
                {
                    windowName = "fuck you";
                }

                Console.WriteLine($"Window name: {windowName}");

                if (lastWindowName != windowName && !string.IsNullOrEmpty(windowName))
                {
                    UpdateStatus(windowName.Substring(0, Math.Min(windowName.Length, 60)));
                    lastWindowName = windowName;
                }

                await Task.Delay(2500);
            }
        }

        private static async void UpdateStatus(string windowName)
        {
            var client = new SlackWebApiClient(Token);
            var response = await client.Users.Profile.Set(new UserProfileSetRequest
            {
                Profile = new UserProfileSet
                {
                    StatusEmoji = ":window:",
                    StatusText = windowName,
                },
            });

            Console.WriteLine($"Request status update to {windowName} Response: {response.OK} Error: {response.Error}");
        }
    }
}