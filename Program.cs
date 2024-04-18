using System;
using System.Windows.Forms;
using System.Drawing;

namespace Macro
{
        public partial class Form1
        {

        }
        static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 form = new Form1(); // Create an instance of Form1
            form.InitializeForm(); // Initialize the form
            Application.Run(form); // Run the application with the initialized form
        }
    }
}
