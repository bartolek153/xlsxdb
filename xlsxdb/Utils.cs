using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Toolkit.Uwp.Notifications;

using Spectre.Console;

namespace xlsxdb
{
    // Log level implementation
    // to use with AnsiConsole
    enum AnsiLogLevel
    {
        Info = 0,
        Warn = 1,
        Error = 2,
        Fatal = 3,
    }

    internal static class Utils
    {
        // Sends a desktop notification (currently, only Windows support)
        public static void Notify(string title, string message)
        {
            // TODO: conditional compilation according to OS?

            new ToastContentBuilder()
                    .AddText(title)
                    .AddText(message)
                    .Show();
        }

        // Uses Spectre.Console to generate log markups.
        public static void Log(AnsiLogLevel level, string message)
        {
            //AnsiConsole.Markup("[green]This is all green[/]");
        }

        // Returns a string array of file paths in the current directory,
        // according to the file extension provided.
        public static string[] GetFilesByExtension(string ext)
        {
            return Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");  // TODO: add support to different Excel filetypes
        }

        // Returns a string array of file paths in the current directory,
        // according to the string array of file extensions provided.
        public static string[] GetFilesByExtension(string[] exts)
        {
            return Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");  // TODO: add support to different Excel filetypes
        }
    }
}
