using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Internkalkyl
{
    static class Logg
    {
      
        static DateTime now = new DateTime();

        public static void toLog(string Msg)
        {
            StreamWriter loggstream = null;
            now = DateTime.Now;

            try
            {
                loggstream = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).ToString() + "\\InternkalkLogg.txt", true);
                loggstream.WriteLine(now.ToLocalTime() + "   " + Msg);
            }
            finally
            {
                if (loggstream != null)
                {
                    loggstream.Close();
                    loggstream.Dispose();
                }
            }            
        }
    }
}
