///
/// Outlook Attachment Extractor 
/// Version 0.1
/// Build 2018-May-11
/// Written by Antoine Dombrecht
/// https://github.com/AntoineDombrecht/SapioCapa.git

using System;
using WinSCP;

namespace MailConnector
{
    class MailConnector
    {
        static void Main(string[] args)
        {
            // OUT
            Output.Initialize();
            // IN
            Input.ReadFile();
            // Save everything in the FTP
            Output.saveInFTP();
            // Clean the temporary folder
            Output.cleanTempDir();
            // END
            Output.End();
        }
    }
}