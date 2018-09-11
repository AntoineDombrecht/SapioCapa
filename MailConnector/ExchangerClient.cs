using System;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace MailConnector
{
    public class ExchangerClient
    {
        //
        private string _userEmailAddress;
        private string _userPassword;
        //
        private ExchangeService service = new ExchangeService();
        private FindItemsResults<Item> listMail;
        private string repTemp;
        //
        private string sender;
        private string subject;
        private string recipient;
        private string recipientType;
        //
        private static string dateTimeFormat = System.Configuration.ConfigurationManager.AppSettings["Program_DateTimeFormat"];
        public ExchangerClient(string _sender, string _subject, string _recipient, string _recipientType)
        {
            Output.WriteInFile("[OK]", "Ouverture de la session Exchanger");
            // Setup temporary file 
            repTemp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
                System.Configuration.ConfigurationManager.AppSettings["Local_TempPath"];
            // Setup main parameters
            _userEmailAddress = System.Configuration.ConfigurationManager.AppSettings["Mail_UserName"];
            _userPassword = System.Configuration.ConfigurationManager.AppSettings["Mail_Password"];
            // Setup credentials
            service.Credentials = new NetworkCredential(_userEmailAddress, _userPassword);
            // Setup URL
            service.Url = new Uri(System.Configuration.ConfigurationManager.AppSettings["Mail_Server"]);
            // Setup mails parameters
            sender = _sender;
            subject = _subject;
            recipient = _recipient;
            recipientType = _recipientType;
            // Filter on mail with specific Subject and/or specific Sender 
            listMail = service.FindItems(WellKnownFolderName.Inbox, filter(sender, subject), new ItemView(50));
            if (service == null)
                Output.WriteInFile("[OK]", "Echec d'ouverture de la session Exchanger");
        }
  
        public void GetAttachmentsFromEmail()
        {
            // Check the folder's existency
            tempFolderExistency();

            // Update the subject of each message locally.
            foreach (Item item in listMail)
            {
                Output.WriteInFile("[OK]", "Début de la récupération de la pièce jointe");
                // Save attachment in temporary folder
                // Bind to an existing message item and retrieve the attachments collection.
                // This method results in an GetItem call to EWS.
                EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
                // Iterate through the attachments collection and load each attachment.
                foreach (Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        // Load the attachment into a file.
                        // This call results in a GetAttachment call to EWS.
                        Output.WriteInFile("[OK]", fileAttachment.Name);
                        Output.WriteInFile("[OK]", repTemp);
                        String[] substrings = System.Configuration.ConfigurationManager.AppSettings["Program_AttachmentType"].Split(',');
                        foreach (string substring in substrings)
                        {
                            if (attachment.Name.Contains(substring))
                            {
                                fileAttachment.Load(repTemp + fileAttachment.Name);
                                System.IO.File.Move(repTemp + fileAttachment.Name,
                                repTemp + DateTime.Now.ToString(dateTimeFormat) + "_" + fileAttachment.Name);
                                Output.WriteInFile("[OK]", "File attachment name: " + fileAttachment.Name);
                            }
                        }
                    }
                    else // Attachment is an item attachment.
                    {
                        ItemAttachment itemAttachment = attachment as ItemAttachment;
                        // Load attachment into memory and write out the subject.
                        // This does not save the file like it does with a file attachment.
                        // This call results in a GetAttachment call to EWS.
                        itemAttachment.Load();
                        Output.WriteInFile("[OK]", "Item attachment name: " + itemAttachment.Name);
                    }
                }
            }
        }
        public void MoveToEmailFolder()
        {
            // Check the folder's existency
            emailFolderExistency();

            Item itemMoved;
            foreach (Item item in listMail)
            {
                itemMoved = item;
                // As a best practice, limit the properties returned by the Bind method to only those that are required.
                PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);

                // Bind to the existing item by using the ItemId.
                // This method call results in a GetItem call to EWS.
                EmailMessage originalMessage = EmailMessage.Bind(service, item.Id, propSet);

                //
                FolderView view = new FolderView(100);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                view.PropertySet.Add(FolderSchema.DisplayName);
                view.Traversal = FolderTraversal.Deep;
                FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);

                //find specific folder
                foreach (Folder f in findFolderResults)
                {
                    //show folderId of the folder recipient
                    if (f.DisplayName == recipient)
                    {
                        itemMoved.Load();
                        // Move the orignal message into another folder in the mailbox and store the returned item.
                        itemMoved = originalMessage.Copy(f.Id);
                    }
                }
                // Check that the item was copied by binding to the copied email message 
                // and retrieving the new ParentFolderId.
                // This method call results in a GetItem call to EWS.
                EmailMessage copiedMessage = EmailMessage.Bind(service, itemMoved.Id, propSet);
                Output.WriteInFile("[OK]", "An email message with the subject '" + originalMessage.Subject +
                    "' has been copied from the inbox folder to the '" + recipient + "' folder.");
                item.Delete(DeleteMode.HardDelete);
                Output.WriteInFile("[OK]", "... and has been deleted from the inbox folder");
            }
        }
        public void emailFolderExistency()
        {
            // Setup the boolean
            bool folderExist = false;

            //Check if the folder exist
            // Create a view with a page size of 100.
            FolderView view = new FolderView(100);

            // Identify the properties to return in the results set.
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);

            // Return only folders that contain items.
            SearchFilter searchFilter = new SearchFilter.IsGreaterThan(FolderSchema.TotalCount, 0);

            // Unlike FindItem searches, folder searches can be deep traversals.
            view.Traversal = FolderTraversal.Deep;

            // Send the request to search the mailbox and get the results.
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, view);

            // Process each item.
            foreach (Folder myFolder in findFolderResults.Folders)
            {
                if (myFolder.DisplayName == recipient)
                {
                    Output.WriteInFile("[OK]", "The folder exists : " + myFolder.DisplayName);
                    folderExist = true;
                }
            }

            // if the folder doesn't already exist create it
            if (!folderExist)
            {
                Output.WriteInFile("[WAR]", "The folder doesn't exist, creating it ...");
                // create the mail folder
                Folder folder = new Folder(service);
                folder.DisplayName = recipient;
                folder.Save(WellKnownFolderName.MsgFolderRoot);
            }
        }
        public void tempFolderExistency()
        {
            if (!Directory.Exists(repTemp))
            {
                Directory.CreateDirectory(repTemp);
            }
        }
        public SearchFilter.SearchFilterCollection filter(string sender, string subject)
        {
            SearchFilter.SearchFilterCollection filter;
            filter = new SearchFilter.SearchFilterCollection();

            if (sender == "" && subject != "")
            {
                if(System.Configuration.ConfigurationManager.AppSettings["Program_RegExSubject"] =="exact")
                    filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, subject));
                else
                    filter.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subject));
            }
            else if (sender != "" && subject == "")
                filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.From, sender));
            else if (sender != "" && subject != "")
            {
                filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.From, sender));
                if (System.Configuration.ConfigurationManager.AppSettings["Program_RegExSubject"] == "exact")
                    filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, subject));
                else
                    filter.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subject));
            }
            return filter;
        }
        // Move a serie of mails into a specific FTP folder
        public bool isEmpty()
        {
            return (listMail.TotalCount == 0) ? true : false;
        }
        /// ACTIONS
        public void deplacer_vers()
        {
            Output.WriteInFile("[OK]", "Début du déplacement des mails");
            if (recipientType == "dossier_ftp")
            {
                Output.WriteInFile("[OK]", "Début du déplacement des mails dans le fichier temporaire : " + recipient);
                repTemp += recipient + @"\";
                // Get the mails attachments
                GetAttachmentsFromEmail();
            }
            if (recipientType == "dossier_mail")
            {
                Output.WriteInFile("[OK]", "Début du déplacement des mails dans le dossier mail : " + recipient);
                this.MoveToEmailFolder();
            }
        }
        public void marquer_comme_lu()
        {
            Output.WriteInFile("[OK]", " Début du marquage des mails ayant pour objet " + subject + " et pour expediteur " + sender);
            if (service != null)
            {
                // Update the subject of each message locally.
                foreach (EmailMessage message in listMail)
                {
                    Output.WriteInFile("[OK]", "Mail datant du " + message.DateTimeReceived);

                    // Setup message status to read
                    if (!message.IsRead) // check that you don't update and create unneeded traffic
                    {
                        message.IsRead = true; // mark as read
                        message.Update(ConflictResolutionMode.AutoResolve); // persist changes
                    }
                    // Print out confirmation with the last eight characters of the item ID and the email subject.
                    Output.WriteInFile("[OK]", "Marked as read local email message with the subject '" + message.Subject + "'.");
                }
            }
            else
                Output.WriteInFile("[ER]", " Le serveur d'Exchanger a rencontré un problème lors de l'accès à la boite mail");
        }
        public void supprimer()
        {
            Output.WriteInFile("[OK]", " Début de la suppression des mails ayant pour objet " + subject + " et pour expediteur " + sender);
            // Update the subject of each message locally.
            foreach (EmailMessage message in listMail)
            {
                message.Delete(DeleteMode.HardDelete);
                // Print out confirmation with the last eight characters of the item ID and the email subject.
                Output.WriteInFile("[OK]", "Delete local email message with the subject '" + message.Subject + "'.");
            }
        }
        public void transferer_a()
        {
            // Create the addresses that the forwarded email message is to be sent to.
            EmailAddress[] addresses = new EmailAddress[1];
            addresses[0] = new EmailAddress(recipient);
            // Create the prefixed content to add to the forwarded message body.
            string messageBodyPrefix = "This is message that was forwarded by using the EWS Managed API";

            if (service != null)
            {
                // Update the subject of each message locally.
                foreach (EmailMessage message in listMail)
                {
                    // Send the forwarded message.
                    message.Forward(messageBodyPrefix, addresses);
                    // Print message subject
                    Output.WriteInFile("[OK]", message.ToRecipients.ToString());
                    // Print out confirmation with the last eight characters of the item ID and the email subject.
                    Output.WriteInFile("[OK]", "Forward the email message with the subject '" + message.Subject + "'.");
                }
            }
            else
                Output.WriteInFile("[ER]", " Le serveur d'Exchanger a rencontré un problème lors de l'accès à la boite mail");
        }
    }
}