using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace glpkXl
{
    class glpkxlMessage
    {
        public double seconds { get; set; }
        public string text { get; set; }
        public glpkxlMessage(double seconds1, string text1)
        {
            seconds = seconds1;
            text = text1;
        }
    }

    static class glpkxlMessages
    {
        public static List<glpkxlMessage> messages;
        public static void initialize()
        {
            messages = new List<glpkxlMessage>();
            stopWatch.start();
        }
        public static void log(Workbook wb, string text)
        {
            log(text);
            double seconds = stopWatch.secondsNow(1);
            wb.Application.StatusBar = "glpkXl: Elapsed: " + seconds + " seconds. Step: " + text;
        }
        public static void log(string text)
        {
            double seconds = stopWatch.secondsNow(3);
            messages.Add(new glpkxlMessage(seconds, text));
        }
        public static void log(Exception e)
        {
            double seconds = stopWatch.secondsNow(3);
            messages.Add(new glpkxlMessage(seconds, e.Message));
            messages.Add(new glpkxlMessage(seconds, e.StackTrace));
        }
        public static void replace(string oldstr, string newstr)
        {
            foreach (var m in messages) m.text = m.text.Replace(oldstr, newstr);
        }
    }
}
