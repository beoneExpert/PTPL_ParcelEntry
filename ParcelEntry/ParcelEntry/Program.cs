using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParcelEntry
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            ParcelEntry PEntry = new ParcelEntry();
            //Application.Run();
            System.Windows.Forms.Application.Run();
        }
    }
}
