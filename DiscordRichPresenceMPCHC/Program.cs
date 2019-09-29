using System;
using System.Text;
using System.Diagnostics;
using System.Configuration;
using System.Threading;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DiscordRichPresenceMPCHC
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        public delegate bool WindowEnumDelegate(IntPtr hwnd, int lParam);

        [DllImport("user32.dll")]
        public static extern int EnumChildWindows(IntPtr hwnd, WindowEnumDelegate del, int lParam);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hwnd, StringBuilder bld, int size);

        static void Main(string[] args)
        {
            // Get Discord client ID from app config file
            long client_id = Convert.ToInt64(ConfigurationManager.AppSettings["client_id"]);
            Console.WriteLine("Client ID: {0}", client_id);

            // Create new Discord GameSDK instance
            var discord = new Discord.Discord(client_id, (ulong)Discord.CreateFlags.Default);

            // Create new Discord Rich Presence activity
            var activity = new Discord.Activity
            {
                ApplicationId = client_id,
                State = "",
                Details = "Not playing anything",
                Timestamps =
                {
                    Start = 0,
                    End = 0
                },
                Assets =
                {
                    LargeImage = "mpc_hc_icon_classic",
                },
            };

            // Set up Discord GameSDK debug logging
            discord.SetLogHook(Discord.LogLevel.Debug, (level, message) =>
            {
                Console.WriteLine("Discord:{0} - {1}", level, message);
            });

            // Get current activity manager from Discord
            var activityManager = discord.GetActivityManager();


            // Main loop: Detect last active MPC-HC window, update presence and run Discord callbacks
            Process activeMPCHCProcess = null;
            while (true)
            {
                bool activityChanged = false;

                if (activeMPCHCProcess != null && activeMPCHCProcess.HasExited)
                {
                    activeMPCHCProcess = null;
                }

                var activeWindowHandle = GetForegroundWindow();

                var processes = Process.GetProcessesByName("mpc-hc64");
                foreach (var process in processes)
                {
                    if (!string.IsNullOrEmpty(process.MainWindowTitle) && process.MainWindowTitle != "Media Player Classic Home Cinema" && process.MainWindowHandle == activeWindowHandle)
                    {
                        activeMPCHCProcess = process;
                        break;
                    }
                }

                if (activeMPCHCProcess != null)
                {
                    if (activity.State != activeMPCHCProcess.MainWindowTitle)
                    {
                        activity.State = activeMPCHCProcess.MainWindowTitle;
                        activityChanged = true;
                    }

                    int i = 0;
                    EnumChildWindows(activeMPCHCProcess.MainWindowHandle, (hwnd, param) =>
                    {
                        string text = "";
                        if (i == 4 || i == 5)
                        {
                            var bld = new StringBuilder(32);
                            GetWindowText(hwnd, bld, 32);
                            text = bld.ToString();
                        }

                        if (i == 4)
                        {
                            var parts = text.Split(new string[] { " / " }, StringSplitOptions.None);
                            if (parts.Length == 2)
                            {
                                TimeSpan startOffset = TimeSpan.Zero, endOffset = TimeSpan.Zero;
                                try
                                {
                                    startOffset = TimeSpan.ParseExact(parts[0], "g", CultureInfo.CurrentCulture);
                                    endOffset = TimeSpan.ParseExact(parts[1], "g", CultureInfo.CurrentCulture);
                                }
                                catch (OverflowException)
                                {
                                    startOffset = TimeSpan.ParseExact(parts[0], "mm\\:ss", CultureInfo.CurrentCulture);
                                    endOffset = TimeSpan.ParseExact(parts[1], "mm\\:ss", CultureInfo.CurrentCulture);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex);
                                }
                                finally
                                {
                                    var now = new DateTimeOffset(DateTime.Now);
                                    var start = now.Subtract(startOffset);
                                    var end = start.Add(endOffset);
                                    var tsStart = start.ToUnixTimeSeconds();
                                    var tsEnd = end.ToUnixTimeSeconds();

                                    if (activity.Timestamps.Start != tsStart)
                                    {
                                        activity.Timestamps.Start = tsStart;
                                        activityChanged = true;
                                    }
                                    if (activity.Timestamps.End != tsEnd)
                                    {
                                        activity.Timestamps.End = tsEnd;
                                        activityChanged = true;
                                    }
                                }
                            }
                        }
                        else if (i == 5)
                        {
                            if (activity.Details != text)
                            {
                                activity.Details = text;
                                activityChanged = true;
                            }
                        }

                        i++;
                        return i <= 5;
                    }, 0);
                }
                else
                {
                    if (activity.State != "")
                    {
                        activity.State = "";
                        activityChanged = true;
                    }
                    if (activity.Details != "Not playing anything")
                    {
                        activity.Details = "Not playing anything";
                        activityChanged = true;
                    }
                    if (activity.Timestamps.Start != 0)
                    {
                        activity.Timestamps.Start = 0;
                        activityChanged = true;
                    }
                    if (activity.Timestamps.End != 0)
                    {
                        activity.Timestamps.End = 0;
                        activityChanged = true;
                    }
                }

                // Update presence if changed
                if (activityChanged)
                {
                    activityManager.UpdateActivity(activity, (result) => {});
                }

                // Run Discord callbacks
                discord.RunCallbacks();

                Thread.Sleep(1000);
            }
        }
    }
}
