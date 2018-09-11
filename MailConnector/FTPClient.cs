using System;
using System.IO;
using WinSCP;

namespace MailConnector
{
    public class FTPClient
    {
        //
        private string userName;
        private string password;
        private string hostName;
        private int port;
        private string protocol;
        private string sshHostKeyFingerprint;
        //
        private SessionOptions sessionOptions = new SessionOptions { };
        private Session session = new Session();

        public FTPClient()
        {
            Output.WriteInFile("[OK]", "Ouverture de la session FTP");
            // Setup main parameters
            userName = System.Configuration.ConfigurationManager.AppSettings["FTP_UserName"];
            password = System.Configuration.ConfigurationManager.AppSettings["FTP_Password"];
            hostName = System.Configuration.ConfigurationManager.AppSettings["FTP_HostName"];
            port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["FTP_Port"]);
            protocol = System.Configuration.ConfigurationManager.AppSettings["FTP_Protocol"];
            sshHostKeyFingerprint = System.Configuration.ConfigurationManager.AppSettings["FTP_SshHostKeyFingerprint"];
            // Setup session options
            sessionOptions.Protocol = Protocol.Ftp;
            sessionOptions.HostName = hostName;
            sessionOptions.PortNumber = port;
            sessionOptions.UserName = userName;
            sessionOptions.Password = password;
            if (protocol == "sftp")
            {
                sessionOptions.Protocol = Protocol.Sftp;
                sessionOptions.SshHostKeyFingerprint = sshHostKeyFingerprint;
            }
            // Setup session log/executable path
            session.ExecutablePath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + 
                @System.Configuration.ConfigurationManager.AppSettings["WinSCP_ExecutablePath"];
            session.DebugLogLevel = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["WinSCP_DebugLogLevel"]);
            session.DebugLogPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + 
                @System.Configuration.ConfigurationManager.AppSettings["WinSCP_DebugLogPath"];
            // Connect
            if(session != null)
                session.Open(sessionOptions);
            else
                Output.WriteInFile("[OK]", "Echec d'ouverture de la session FTP");
        }
    
        public void saveInFile(string recipient)
        {
            string repTemp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
                System.Configuration.ConfigurationManager.AppSettings["Local_TempPath"];
            repTemp+= recipient + @"\";
            // Check the folder's existency
            if (!session.FileExists("/" + recipient))
            {
                Output.WriteInFile("[WAR]", "The folder doesn't exist in the FTP, creating it ...");
                session.CreateDirectory("/" + recipient);
            }

            Output.WriteInFile("[OK]", "Sauvegarde dans le fichier FTP à partir de : "+ recipient);
            // Upload files
            TransferOptions transferOptions = new TransferOptions();
            transferOptions.TransferMode = TransferMode.Binary;
            TransferOperationResult transferResult;
            transferResult = session.PutFiles(repTemp.Replace("/", @"\") + @"*", "/" + recipient + "/", false, transferOptions);
            // Throw on any error
            transferResult.Check();
            // Print results
            foreach (TransferEventArgs transfer in transferResult.Transfers)
            {
                Output.WriteInFile("[OK]", "Upload of " + transfer.FileName + " succeeded");
            }
        }
    }
  }