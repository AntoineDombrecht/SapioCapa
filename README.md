# SapioCapa
Gestionnaire intelligent de capacité

Prérequis :

Ce logiciel fonctionne avec la version 4.0 de la librairie .NET, toute version anterieure à celle-ci pourrait entrainer des dysfonctionnements.

Afin de pouvoir décrypter les données sensibles utilisées pour la configuration du programme (mots de passe, noms de compte utilisateur et identifiants FTP) le logiciel ASP_Regiis.exe est utilisé. https://msdn.microsoft.com/fr-fr/library/k6h9cz8h(v=vs.100).aspx
Celui-ci est présent dans le répertoire C:\WINDOWS\Microsoft.NET\Framework\[v4.0.30319]\

Informations sur la configuration :

Les données de configuration sont situées dans le fichier >"MailConnector.exe.config" pour le décrypter veillez utiliser la commande suivante : "rename MailConnector.exe.config web.config && "[chemin vers regiis]\aspnet_regiis.exe" -pdf "appSettings" "[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config"

La commande de cryptage est quasi similaire : >"rename MailConnector.exe.config web.config && "[chemin vers regiis]\aspnet_regiis.exe" -pef "appSettings" "[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config"

On retrouve dans le fichier de configuration les données concernant le serveur mail, le serveur ftp, les chemins relatifs et le programme.

Le serveur FTP
→ Le champ 'FTP_Protocol' correspond au type du serveur, il peut prendre les valeurs suivantes : 'ftp' ou 'sftp'
→ Le champ 'FTP_HostName', l'adresse du serveur 
→ Le champ ‘FTP_Port’, le port du serveur
→ Le champ ‘FTP_UserName’, l'adresse de connexion au FTP
→ Le champ ‘FTP_Password’, le mot de passe
→ Le champ ‘FTP_SshHostKeyFingerprint’ value=‘ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:...’ 
→ Le champ ‘Mail_UserName’ value=‘antoine.dombrecht@thalesgroup.com’ 
→ Le champ ‘Mail_Password’ value=‘/AlcdmN21011997’ 
→ Le champ ‘Mail_Server’ value=‘https://email.iris.infra.thales/EWS/Exchange.asmx’ 
→ Le champ ‘WinSCP_ExecutablePath’ value=‘\WinSCP.exe’ 
→ Le champ ‘WinSCP_DebugLogPath’ value=‘\winscp.log’ 
→ Le champ ‘Local_TempPath’ value=‘\temp\’ 
→ Le champ ‘Local_XMLPath’ value=‘\data-set.xml’  
→ Le champ ‘Local_XLSXPath’ value=‘\data-set.xlsx’ 
→ Le champ ‘Local_DebugPath’ value=‘\output.log’ 
→ Le champ ‘Program_Col1’ value=‘1’ 
‘Program_Col2’ value=‘2’ 
‘Program_Col3’ value=‘3’ 
‘Program_Col4’ value=‘4’ 
‘Program_Col5’ value=‘5’ 
‘Program_Col6’ value=‘7’ 
‘Program_Col7’ value=‘9’ 



Informations sur le paramétrage :

Une fois la configuration faite vous pouvez commencer à paramétrer le logiciel. Le paramétrage peut se faire, soit dans le fichier input-data.xlsx soit dans le fichier input-data.xml. Si aucun fichier excel n'est spécifié dans la configuration du programme ou bien si le fichier n'existe pas, seul le fichier XML sera lu. Dans le cas contraire, le programme lira le fichier excel et générera le xml correspondant.


/!\ Précautions à prendre /!\
