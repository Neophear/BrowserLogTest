using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BrowserLogTest
{
    //http://www.codeproject.com/Articles/12105/Internet-Explorer-Activity-Monitor

    class Program
    {
        private static List<IEContainer> Current_IEs;
        private static SHDocVw.ShellWindows shellWindows;

        static void Main(string[] args)
        {
            Current_IEs = new List<IEContainer>();
   		    shellWindows = new SHDocVw.ShellWindowsClass();
   		    shellWindows.WindowRegistered+=new SHDocVw.DShellWindowsEvents_WindowRegisteredEventHandler(shellWindows_WindowRegistered);
   		    shellWindows.WindowRevoked+=new SHDocVw.DShellWindowsEvents_WindowRevokedEventHandler(shellWindows_WindowRevoked);
            
            while (true)
                Console.ReadLine();
        }

        // THIS EVENT IS FIRED WHEN THE A NEW BROWSER IS CLOSED
  	    private static void shellWindows_WindowRevoked(int z)
  	    {
            IEContainer iec = Current_IEs.Find(x => x.Cookie == z);

            if (iec != null)
            {
                Console.WriteLine("Browser close: " + z);
                //iec.IE.NavigateComplete2 -= new SHDocVw.DWebBrowserEvents2_NavigateComplete2EventHandler(browser_NavigateComplete2);
                //while (Marshal.ReleaseComObject(iec.IE) > 0)
                //{

                //}
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(iec.IE);
                //iec.IE = null;
                Current_IEs.Remove(iec);
            }
  	    }
  
  	    // THIS EVENT IS FIRED WHEN THE A NEW BROWSER IS OPEN
        private static void shellWindows_WindowRegistered(int z)
        {
            string filnam;

            foreach (SHDocVw.InternetExplorer ie in shellWindows)
            {
                filnam = Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
                
                if (filnam.Equals("iexplore"))
                {
                    if (!Current_IEs.Exists(x => x.IE == ie))
                    {
                        try
                        {
                            bool adbar = (bool)ie.GetType().InvokeMember("AddressBar", System.Reflection.BindingFlags.GetProperty, null, ie, null);

                            if (adbar)
                            {
                                Current_IEs.Add(new IEContainer { Cookie = z, IE = ie });
                                Console.WriteLine("Browser open: " + z);
                                ie.NavigateComplete2 += browser_NavigateComplete2;
                            }
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
            }
        }

        static void browser_NavigateComplete2(object pDisp, ref object URL)
        {
            try
            {
                bool adbar = (bool)pDisp.GetType().InvokeMember("AddressBar", System.Reflection.BindingFlags.GetProperty, null, pDisp, null);

                if (adbar)
                    Console.WriteLine(URL);
            }
            catch (Exception)
            {

            }
        }

        class IEContainer : IEquatable<IEContainer>
        {
            public int Cookie { get; set; }
            public SHDocVw.InternetExplorer IE { get; set; }

            public bool Equals(IEContainer other)
            {
                return other != null && other.Cookie == this.Cookie;
            }
            public override bool Equals(object obj)
            {
                if (obj == null)
                    return false;

                IEContainer p = obj as IEContainer;

                if (p == null)
                    return false;

                return (Cookie == p.Cookie);
            }
            public override int GetHashCode()
            {
                return Cookie;
            }
        }
    }
}