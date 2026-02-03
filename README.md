# SapioCapa - MailConnector STS

Gestionnaire intelligent de capacité, gestion des données stockées sur Microsoft Exchange.

### Pour commencer

1. Décompresser le fichier d'installation.
2. Installer [WinSCP](https://winscp.net/eng/download.php).
3. Remplir le fichier de configuration (se reporter à la section *informations sur la configuration*).
4. Lancer l'executable.
5. Enjoy :+1:

Optionnel : Si vous souhaitez chiffrer ou déchiffrer le fichier de configuration, reportez vous à la section *information sur la configuration*. <br />
 <br />
Indication : La trace d'exécution du programme peut être trouvée dans *output.log* (par défaut).

### Prérequis

Ce logiciel fonctionne avec la version 4.0 de la librairie .NET, toute version antérieure à celle-ci pourrait entrainer des disfonctionnements. <br />
 <br />
Afin de pouvoir déchiffrer les données sensibles utilisées pour la configuration du programme (mots de passe, noms de compte utilisateur et identifiants FTP) le logiciel [ASP_Regiis.exe](https://msdn.microsoft.com/fr-fr/library/k6h9cz8h(v=vs.100).aspx) est utilisé. 
Celui-ci est présent dans le répertoire *C:\WINDOWS\Microsoft.NET\Framework\[v4.0.30319]\\*. <br />
 <br />
Ce logiciel pilote [WinSCP](https://winscp.net/eng/download.php), il faut donc indiquer au programme où se situe l'executable WinSCP comme indiqué ci-après. <br />

### Informations sur la configuration

Les données de configuration sont situées dans le fichier *MailConnector.exe.config* pour le déchiffrer veillez utiliser la commande suivante : ```rename MailConnector.exe.config web.config && "\[chemin vers regiis]\aspnet_regiis.exe" -pdf "appSettings" "\[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config``` <br />
 <br />
La commande de chiffrement est quasi similaire : ```rename MailConnector.exe.config web.config && "\[chemin vers regiis]\aspnet_regiis.exe" -pef "appSettings" "\[chemin vers le repertoire d'installation]\bin\release" && rename web.config MailConnector.exe.config``` <br />
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
* **Mail_UserName** - l'adresse mail de l'utilisateur
* **Mail_Password** - le mot de passe de l'utilisateur
* **Mail_Server** - l'adresse du serveur

#### Les chemins
* **WinSCP_ExecutablePath** - le chemin vers l'executable WinSCP (pilotage de WinSCP) 
* **WinSCP_DebugLogPath** - le chemin vers les log WinSCP
* **Local_TempPath** - le chemin vers le dossier cache
* **Local_XMLPath** - le chemin vers le fichier XML (si non présent, il sera généré depuis le fichier excel)
* **Local_XLSXPath** - le chemin vers le fichier Excel (si non présent, le programme lira le fichier XML)
* **Local_DebugPath** - le chemin vers les logs du programme
* **Program_ColX** - le numéro des colonnes lues par le programme dans le fichier Excel
* **Program_DateTimeFormat** - Indique le format d'écriture de la date. Il est utilisé pour la sauvegarde des fichiers dans le FTP, il doit donc être conforme aux règles d'écritures des noms de fichiers (caractères \\/\*:?"<>| interdits).
* **Program_RegExSubject** - Expression régulière concernant l'écriture des objets dans le tableau Excel. Si la valeur est "exact" seuls les mails dont l'objet est strictement celui indiqué dans le tableau seront traités. Dans le cas contraire (i.e. pour toutes les autres valeurs du champ) le programme selectionnera les mails dont l'objet contient l'expression indiquée.
* **Program_AttachmentType** - Types de format pris en charge lors de l'importation des pièces jointes séparés par des virgules et sans espace (e.g ".pdf,.xslx,.csv").
* **Program_TempSize** - Taille en octets au delà de laquelle le fichier de log du programme sera supprimé. Si cette taille vaut 0, le fichier de log sera supprimé après chaque exécution.

### Informations sur le paramétrage

Une fois la configuration faite vous pouvez commencer à paramétrer le logiciel. Le paramétrage peut se faire, soit dans le fichier input-data.xlsx soit dans le fichier input-data.xml. Si aucun fichier excel n'est spécifié dans la configuration du programme ou bien si le fichier n'existe pas, seul le fichier XML sera lu. Dans le cas contraire, le programme lira le fichier excel et générera le xml correspondant.


### /!\ Précautions à prendre /!\

#### Concernant le fonctionnement du programme.

Le programme peut fonctionner avec la structure minimale suivante <br />
.  <br />
├── bin  <br />
│ &nbsp; &nbsp; &nbsp; &nbsp; └── release  <br />
│&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ├── WinSCPnet.dll  <br />
│&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ├── Microsoft.Exchange.WebServices.dll  <br />
│&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ├── MailConnector.exe.config  <br />
│&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; └── MailConnector.exe  <br />
├── data-set.xlsx et/ou data-set.xml <br />
├── WinSCP.exe <br />
└── README.md <br />
<br />
Dans le cas où le fichier XML n'est pas déjà présent, il est généré par le programme depuis le fichier Excel, il en va de même pour les fichiers de logs (winscp.log et output.log).

#### Concernant le fichier excel

Cette version implémente les actions suivantes : ```Marque comme lu, Déplacer vers, transférer à, supprimer```
L'accentuation et la mise en forme des entrées du tableau n'a pas d'importance, seul l'orthographe compte.
Une suite d'actions est possible si on sépare chacune d'entre elle par une virgule sans laisser d'espace, exemple : ```Marquer comme lu,déplacer vers,supprimer```. <br />
On ne peut en revanche indiquer qu'une valeur pour le type de destinataire (ou destination), exemple : ```adresse mail,fichier mail``` n'est pas possible. Par extension la suite d'action ```Transférer à,déplacer vers``` n'est pas possible car elle nécessite de rentrer l'adresse mail et le fichier dans le même champ destinataire.<br />
Veillez à toujours étendre les tableaux plutôt que les diminuer. Si vous le diminuez verifiez que les cellules du tableau ne soient pas verrouillées.<br />
```diff 
- Attention
```
**TOUT DEPLACEMENT EST IRREVERSIBLE** Les mails traités le sont depuis la boite de réception uniquement. Dès lors qu'un mail est déplacé ailleurs, il va de soi qu'il ne pourra plus être pris en compte lors des futures opérations. 
#### Concernant le fichier XML

Un fichier *data-set.xsd* indique la forme que doit prendre le fichier XML afin d'être conforme au parseur du programme. <br/>
Le fichier xsd doit toujours être présent dans le répertoire d'installation, ceci afin que le programme puisse alerter l'utilisateur d'un défaut d'écriture dans le xml.


