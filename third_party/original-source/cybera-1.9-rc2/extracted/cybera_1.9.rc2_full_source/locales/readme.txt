_______________________________________________________________________________
This program is distributed under the terms of the GNU General Public License.

This folder contains the localization sub-projects for Cybera.

The cybcli folder contains the workstation localizations and the cybserv folder
contains the server localization.
In each one of these folders can be found sub-folders for each language.

In order to create a new language project you can copy the French folder (fr)
to say 'es' for Spanish, you then open the new folder and rename the files from
*_fr.* to *_es.*. Should you prefer to work from the English resources, you can
replace the res.res file with the one in the 'en' folder.
You can now open the project file (*_es.vbp) and rename the project and the 
application title in the project properties.

I have started my string table at 10000. All string below 10100 are common to
the server and the client. All strings should be translated EXEPT 10006 as this
is used for path completion and should stay in English. String 10014 is the name
of the language.

The resulting dll needs to be registered in order to be useable. 
Should you be interested in translating Cybera, please send me the new language
project so that it can be incorporated into the next release.

In order to complete a translation you also need to provide the 5 RTF files that
are dsiplayed on the workstation, to the left of the logo. These can be found in
the server options and are named en_blocked.rtf, en_intro.rtf, en_nocredit.rtf,
en_sessionend.rtf, en_sessionpaused.rtf for example.
.

Contact info:
Email: smonsarr@users.sourceforge.net
Web: http://cybera.sourceforge.net/

