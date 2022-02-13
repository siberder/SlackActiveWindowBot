using System;
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

        private static string GetActiveWindowTitle()
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
                var windowName = GetActiveWindowTitle();

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