using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestSpeech
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            // Add the event handler for handling UI thread exceptions to the event.
            System.Windows.Forms.Application.ThreadException += new ThreadExceptionEventHandler(MyCommonExceptionHandlingMethod);

            // Set the unhandled exception mode to force all Windows Forms errors to go through
            // our handler.
            System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            // Add the event handler for handling non-UI thread exceptions to the event. 
            AppDomain.CurrentDomain.UnhandledException +=
                new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void ShowError(string msg)
        {
            MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void MyCommonExceptionHandlingMethod(
                                              object sender,
                                              ThreadExceptionEventArgs t)
        {
            ShowError(t.Exception.Message);  
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ShowError(e.ExceptionObject.ToString());
        }


    }
}
