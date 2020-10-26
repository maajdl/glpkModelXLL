using System;

namespace glpkXl
{
    static class stopWatch
    {
        static DateTime t0;
        static DateTime t1;
        public static void start()
        {
            t0 = DateTime.Now;
        }
        public static void stop()
        {
            t1 = DateTime.Now;
        }
        public static double seconds(int decimals)
        {
            double dt = Math.Round(t1.Subtract(t0).TotalSeconds, decimals);
            return dt;
        }
        public static string seconds(string fmt)
        {
            string dt = t1.Subtract(t0).TotalSeconds.ToString(fmt);
            return dt;
        }
        public static double secondsNow(int decimals)
        {
            double dt = Math.Round(DateTime.Now.Subtract(t0).TotalSeconds, decimals);
            return dt;
        }
        public static string secondsNow(string fmt)
        {
            string dt = DateTime.Now.Subtract(t0).TotalSeconds.ToString(fmt);
            return dt;
        }
    }
}
