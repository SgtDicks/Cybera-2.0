$Id: readme.txt 17 2006-01-04 09:47:44Z sam $
________________________________________________________________________________
English:
This program is distributed under the terms of the GNU General Public License.
Cybera is a cyber café administration system. It is comprised of two programs,
a client installed on each workstation piloted by a server application 
installed on a central server.
Cybera is written in Visual Basic 6. It uses a Jet database for storing cards, 
recharges and session and income history. It is designed to work on Win2000 or 
WinXP Pro.

The source code is comprised of two main projects (cybserv & cybcli sub 
directories) that are managed together as a group (cybadmin - grpCybadmin.vbg)

Source code map:
cybadmin        :   Root source code directory, shared source code
|-cybcli        :   Cybera Workstation source code
|-cybserv       :   Cybera Server source code
|-locales       :   Language pack DLL source code for both Worksation & Server
| |-cybcli      :   Workstation language packs
| | |-br        :   Brazilian Portuguese language pack project
| | |-de        :   German language pack project
| | |-el        :   Greek language pack project
| | |-en        :   English resource file
| | |-es        :   Spanish language pack project
| | |-fr        :   French language pack project
| | |-hu        :   Hungarian language pack project
| | `-it        :   Italian language pack project
| `-cybserv     :   Server language packs
| | |-br        :   Brazilian Portuguese language pack project
| | |-de        :   German language pack project
| | |-el        :   Greek language pack project
| | |-en        :   English resource file
| | |-es        :   Spanish language pack project
| | |-fr        :   French language pack project
| | |-hu        :   Hungarian language pack project
| | `-it        :   Italian language pack project
|-resources     :   Resources used in Cybera (all the icons)
`-Setup         :   Nullsoft installer scripts & resources
  |-resources   :   Cybera Resources included in the installers
  |-scripts     :   Server & workstation installer scripts
  `-system      :   VB runtime DLLs & OCXs plus 3rd party DLLs

NOTE: Cybera cannot be run from the source code archive only. The Cybera 
      executables will not work correctly if the installer has not been executed
	  to create the required file environment.
________________________________________________________________________________
Français :
Ce programme est distribué sous les termes de la Licence Publique Générale GNU.
Cybera est un système de gestion de cyber café. Il fonctionne sur le modèle 
client serveur où un ordinateur : le serveur, pilote et contrôle un ensemble 
d'autres ordinateurs : les postes de travail. 
Cybera est écrit en Visual Basic 6. Il utilise une base de données Jet pour 
stocker les cartes,clients et l'historique d'utilisation. Ce programme est 
conçu pour win2000 ou winXP Pro

Le code source est composé de deux projets principaux (les sous-dossiers 
cybserv & cybcli) qui sont réunis dans un groupe (cybadmin - grpCybadmin.vbg)

Plan du code source :
cybadmin        :   Dossier racine du code, code partagé
|-cybcli        :   Code du Client Cybera
|-cybserv       :   Code du Serveur Cybera 
|-locales       :   Code des DLLs des packs de langue pour Serveur & Client
| |-cybcli      :   Packs de langue du Client Cybera
| | |-br        :   Projet de pack de langue  Portugais brésilien
| | |-de        :   Projet de pack de langue Allemand
| | |-el        :   Projet de pack de langue Grèque
|   |-en        :   Fichier resource anglais
| | |-es        :   Projet de pack de langue Espagnol
| | |-fr        :   Projet de pack de langue Français
| | |-hu        :   Projet de pack de langue Hongrois
| | `-it        :   Projet de pack de langue Italien
| `-cybserv     :   Packs de langue du Serveur Cybera
| | |-br        :   Projet de pack de langue  Portugais brésilien
| | |-de        :   Projet de pack de langue Allemand
| | |-el        :   Projet de pack de langue Grèque
|   |-en        :   Fichier resource anglais
| | |-es        :   Projet de pack de langue Espagnol
| | |-fr        :   Projet de pack de langue Français
| | |-hu        :   Projet de pack de langue Hongrois
| | `-it        :   Projet de pack de langue Italien
|-resources     :   Resources utilisées par Cybera (les icônes)
`-Setup         :   Scripts & resources des installeurs Nullsoft 
  |-resources   :   Resources Cybera inclus dans les installeurs
  |-scripts     :   Scripts des installeurs de Serveur & Client Cybera
  `-system      :   DLLs & OCXs VB plus les DLLs tierces

NOTE : Cybera ne pent pas être demarré à partir de l'archive du code source
       uniquement. Les executables Cybera ne marcheront que si les instalateurs
	   ont été éxecutés afin de créer l'environement fichier requis.

Contact info:
Email: smonsarr@users.sf.net
Web: http://cybera.sf.net/

