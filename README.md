# SapioCapa - MailConnector

Gestionnaire intelligent de capacité, gestion des mails.

### Pour commencer

1. Décompresser le fichier d'installation.
2. Installer [WinSCP](https://winscp.net/eng/download.php).
3. Remplir le fichier de configuration (se reporter à la section *informations sur la configuration*).
4. Lancer l'executable.
 <br />
Optionnel : Si vous souhaitez crypter ou décrypter le fichier de configuration, reportez vous à la section *information sur la configuration*. <br />
 <br />
Indication : La trace d'exécution du programme peut être trouvée dans *output.log* (par défaut).

### Prérequis

Ce logiciel fonctionne avec la version 4.0 de la librairie .NET, toutes versions anterieures à celle-ci pourraient entrainer des dysfonctionnements. <br />
 <br />
Afin de pouvoir décrypter les données sensibles utilisées pour la configuration du programme (mots de passe, noms de compte utilisateur et identifiants FTP) le logiciel [ASP_Regiis.exe](https://msdn.microsoft.com/fr-fr/library/k6h9cz8h(v=vs.100).aspx) est utilisé. 
Celui-ci est présent dans le répertoire *C:\WINDOWS\Microsoft.NET\Framework\[v4.0.30319]\*. <br />
 <br />
Ce logiciel pilote [WinSCP](https://winscp.net/eng/download.php), il faut donc indiquer au programme où se situe l'executable WinSCP comme indiqué ci-après. <br />

### Informations sur la configuration

Les données de configuration sont situées dans le fichier *MailConnector.exe.config* pour le décrypter veillez utiliser la commande suivante : ```rename MailConnector.exe.config web.config && "\[chemin vers regiis]\aspnet_regiis.exe" -pdf "appSettings" "\[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config``` <br />
 <br />
La commande de cryptage est quasi similaire : ```rename MailConnector.exe.config web.config && "\[chemin vers regiis]\aspnet_regiis.exe" -pef "appSettings" "\[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config``` <br />
 <br />
On retrouve dans le fichier de configuration les données concernant le serveur mail, le serveur ftp, les chemins relatifs et le programme.

#### Le serveur FTP
* **‘FTP_Protocol’**  - correspond au type du serveur, il peut prendre les valeurs suivantes : 'ftp' ou 'sftp'
* **‘FTP_HostName’** - l'adresse du serveur 
* **‘FTP_Port’** - le port du serveur
* **‘FTP_UserName’** - l'adresse de connexion au FTP
* **‘FTP_Password’** - le mot de passe de connexion au FTP
* **‘FTP_SshHostKeyFingerprint’** - la clé SSH dans le cas d'une connexion sécurisée au FTP 

#### Le serveur Mail
* **‘Mail_UserName’** - l'adresse mail de l'utilisateur
* **‘Mail_Password’** - le mot de passe de l'utilisateur
* **‘Mail_Server’** - l'adresse du serveur

#### Les chemins
* **‘WinSCP_ExecutablePath’** - le chemin vers l'executable WinSCP (pilotage de WinSCP) 
* **‘WinSCP_DebugLogPath’** - le chemin vers les log WinSCP
* **‘Local_TempPath’** - le chemin vers le dossier cache
* **‘Local_XMLPath - le chemin vers le fichier XML (si non présent, il sera généré depuis le fichier excel)
* **‘Local_XLSXPath’** - le chemin vers le fichier Excel (si non présent, le programme lira le fichier XML)
* **‘Local_DebugPath’** - le chemin vers les logs du programme
* **‘Program_ColX’** - le numéro des colonnes lues par le programme dans le fichier Excel
* **‘Program_DateTimeFormat’** - Indique le format d'écriture de la date. Il est utilisé pour la sauvegarde des fichiers dans le FTP, il doit donc être conforme aux règles d'écritures des noms de fichiers (caractères \\/\*:?"<>| interdits).

### Informations sur le paramétrage

Une fois la configuration faite vous pouvez commencer à paramétrer le logiciel. Le paramétrage peut se faire, soit dans le fichier input-data.xlsx soit dans le fichier input-data.xml. Si aucun fichier excel n'est spécifié dans la configuration du programme ou bien si le fichier n'existe pas, seul le fichier XML sera lu. Dans le cas contraire, le programme lira le fichier excel et générera le xml correspondant.


### /!\ Précautions à prendre /!\

#### Concernant le fonctionnement du programme.

Le programme peut fonctionner avec la structure minimale suivante <br />
|- bin - release | - WinSCPnet.dll <br />
                 | - Microsoft.Exchange.WebServices.dll <br />
                 | - MailConnector.exe.config <br />
                 | - MailConnector.exe <br />
|- data-set.xlsx et/ou data-set.xml <br />
|- WinSCP.exe <br />
 <br />
Dans le cas où le fichier XML n'est pas déjà présent, il est généré par le programme depuis le fichier Excel, il en va de même pour les fichiers de logs (winscp.log et output.log).

#### Concernant le fichier excel

Cette version implémente les actions suivantes : ```Marque comme lu, Déplacer vers, transéferer à, supprimer```
L'accentuation et la mise en forme des entrées du tableau n'a pas d'importance, seul l'orthographe compte.
Une suite d'actions est possible si on sépare chacune d'entre elle par une virgule sans laisser d'espace, exemple : ```Marquer comme lu,déplacer vers,supprimer```. <br />
On ne peut en revanche indiquer qu'une valeur pour le type de destinataire (ou destination), exemple : ```adresse mail,fichier mail``` n'est pas possible. Par extension la suite d'action ```Transférer à,déplacer vers``` n'est pas possible car elle nécessite de rentrer l'adresse mail et le fichier dans le même champ destinataire.

#### Concernant le fichier XML

Un fichier *data-set.xsd* indique la forme que doit prendre le fichier XML afin d'être conforme au parseur du programme.


