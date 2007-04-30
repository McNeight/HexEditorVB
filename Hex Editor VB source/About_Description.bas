Attribute VB_Name = "About_Description"

' =======================================================
' Hex Editor VB for Windows
' Copyright (c) 2006-2007 Alain Descotes (violent_ken)
' https://sourceforge.net/projects/hexeditorvb/
' =======================================================




 


' =======================================================
' LICENSE (FRANCAIS)
' =======================================================

' Ce logiciel est sous license GNU General Public License. La description officielle de la
' licence n'est pas disponible en fran�ais, mais vous pouvez trouver une traduction (non
' officielle) ici : http://fsffrance.org/gpl/gpl-fr.fr.html
'
' Ce logiciel n'est distribu� sans AUCUNE GARANTIE (se reporter � la licence pour les d�tails)
'
' Vous devriez aoir re�u une copie de la licence avec ce code ou ce logiciel, sinon �crite
' � la Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston,
' MA  02111-1307  USA




' =======================================================
' LICENSE (ENGLISH)
' =======================================================

' Hex Editor VB is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' Hex Editor VB is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with Hex Editor VB; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA







' =======================================================
' DESCRIPTION DE HEX EDITOR VB (FRANCAIS)
' =======================================================
'
' Hex Editor VB est un �diteur hexad�cimal pour Windows XP/Vista.
' Utilisation confortable avec une r�solution minimale de 1024*768. Pour disposer de toutes
' les fonctionnalit�s du logiciel, les droits d'administrateur sont requis
' (notamment l'ouverture de processus et l'�criture dans le registre).
'
' Il inclut les fonctionnalit�s standarts d'�dition de fichier : supporte des fichiers de
' taille presque > 900 To, copier/coller/supprimer (permet de raccourcir les
' fichiers), ins�rer (permet d'augmenter la taille), historique (pour r�tablir le fichier
' dans son �tat d'origine), cr�ation de fichier depuis la s�lection, recherche de strings
' ou de valeurs hexa, remplacement de strings...
'
' Inclut �galement l'�dition des 2Go de m�moire virtuelle de chaque processus du syst�me
' (lecture et �criture pour les zones m�moires non prot�g�es par le syst�me). Fonctions
' classiques similaire � l'�dition de fichier.
'
' Permet �galement l'�dition des disques durs physiques ou des partition logiques
' (support des formats FAT16, FAT32, NTFS, CDFS et UDF). Permet les fonctions
' classiques similaires � l'�dition de fichier, avec en plus la gestion de la carte des
' clusters de chaque fichier.
'
' Inclut �galement des outils de gestion des fichiers (renommage massif, suppression
' s�curis�e des fichiers, comparaisons, changement de dates, r�cup�ration des fichiers,
' d�coupage/fusion, recherche de fichiers/contenu...)
' et des processus (gestionnaire de taches int�gr�).
'
' Inclut aussi en �diteur de script pour automatiser les taches r�currentes.
'
' Inclut �galement la possibilit� de convertir des donn�es de diff�rents types entre eux
' (binaire, hexad�cimal, d�cimal, octal, ANSI ASCII, base n).
'
' Le logiciel permet la cr�ation/sauvegarder/chargement de signets, la personnalisation de
' la visualisation du fichier, l'impl�mentation dans le menu contextuel de Explorer,
' un explorateur de fichiers pour permettre l'ouverture rapide, un explorateur de disque...
'
' Inclut �galement un module de d�sassemblage d'ex�cutables Win32.




' =======================================================
' DESCRIPTION OF HEX EDITOR VB (ENGLISH)
' =======================================================
'
' Hex Editor VB is a hexadecimal editor for Windows XP/Vista.
' Comfortable use with a minimal resolution of 1024*768. To have all the functionalities
' of the software, the rights of administrator are necessary (in particular the opening of
' process and the writing in the registry).
'
' It includes the standarts functionalities  of file edition : it supports files of
' size > 900 To, copy/paste/remove (allows to reduce file size), insertion (allows to
' increase the size), history (to restore the file in its state of origin), creation of file
' from the selection, search of strings or hexa values, replacement of strings�
'
' Also support of the edition of the 2Go of virtual memory of each process of the system
' (reading and writing for the zones wich are not protected by the system).
' Traditional functions similar to the edition of file.
'
' Also support of the edition of the physical hard disks or logical partition (support of
' FAT16, FAT32, NTFS, CDFS and UDF). Allows the traditional functions similar to the file
' edition with in more management of the chart of the clusters of each file.
'
' Also includes management tools for files (massive renaming, definitive suppression of
' files, comparisons, change of dates, recovery of files, cutting/fusion, search for
' files/contained�) and for processes (task manager integrated).
'
' Also includes an editor of script to automate the recurring spots.
'
' Also includes the possibility of converting data of various types between them (binary,
' hexadecimal, decimal, octal, ASCII ANSI, n-bases).
'
' The software allows creation/saving/loading of bookmarks, the personalization of the
' visualization of the file, the implementation in the contextual menu of Explorer,
' a file explorer to allow the fast opening, a disk explorer�
'
' Also includes a disassembler module for Win32 binaries.






' =======================================================
' CONTACTS & LIENS (FRANCAIS)
' =======================================================
'
' Vous pouvez poser vos questions sur le forum de sourceforge.net d�di� � mon projet :
' https://sourceforge.net/forum/?group_id=186829
'
' Vous pouvez aussi me contacter � l'adresse hexeditorvb@gmail.com
'
' La page de t�l�chargement du projet :
' https://sourceforge.net/project/showfiles.php?group_id=186829
'
' Vous pouvez me faire parvenir les bugs non r�pertori�s � cette adresse :
' https://sourceforge.net/tracker/?group_id=186829
'
' La page principale du projet :
' https://sourceforge.net/projects/hexeditorvb/
'
' Le site Internet h�berg� par sourceforge.net :
' http://hexeditorvb.sourceforge.net/
'
' SVP pensez � me faire parvenir vos logs contenant la liste des bugs rencontr�s
' (Aide --> Rapports d'erreurs)




' =======================================================
' CONTACTS & LINKS (ENGLISH)
' =======================================================
'
' Please ask your question about Hex Editor VB here :
' https://sourceforge.net/forum/?group_id=186829
'
' You can also contact me : hexeditorvb@gmail.com
'
' To download the lastest versions :
' https://sourceforge.net/project/showfiles.php?group_id=186829
'
' You can report bugs here :
' https://sourceforge.net/tracker/?group_id=186829
'
' Principal page of Hex Editor VB on sourceforge.net :
' https://sourceforge.net/projects/hexeditorvb/
'
' Website :
' http://hexeditorvb.sourceforge.net/
'
' Please send me your log files with the list of bugs (Help --> Bug report)







' =======================================================
' HISTORIQUE DES VERSIONS (FRANCAIS)
' =======================================================
'
' VERSIONS FINALES :
'
'
' VERSIONS BETA :
'
'
' VERSIONS ALPHA :
'
'
' VERSIONS PRE ALPHA :
'
'   v1.7
'   -Ajout d'un outil de cr�ation de fichiers ISO
'   -Ajout d'un outil de sanitization des disques/fichiers
'
'   v1.6
'   -Optimisations diverses du logiciel (d�marrage, ouverture des fichiers..)
'   -Ajout de la gestion de fin de fichier
'   -Ajout de la console
'   -Ajout de nouvelles options,
'   -Nombreux bugs corrig�s
'   -Ajout� le Disassembler avec le support multilingue et les traductions fran�aise et anglaise compl�tes
'   -Ajout de la s�lection et du support des disques physiques
'   -Bugs d'affichage r�solus
'   -Optimis� la vitesse de d�marrage de l'explorateur de fichiers
'   -Ajout d'un gestionnaire de signets
'   -Ajout de l'exportation vers le presse papier et vers un fichier pour l'exportation de fichiers complets
'   -Ajout d'arborescences de processus plut�t que de simples listes
'   -Ajout du projet "Editeur de langue"
'   -Ajout du support multi-langue complet pour Hex Editor VB except� pour la console
'   -Ajout de la traduction fran�aise compl�te
'
'   v1.5
'   -Ajout de la recherche de fichiers
'   -Ajout des options
'   -Ajout des icones
'   -Correction de nombreux bugs
'   -Support des formats CDFS et UDF
'   -Support des fichiers de plus de 900To
'
'   v1.4
'   -Ajout des options
'   -Ajout de la recherche en m�moire
'   -Ajout de la conversion avanc�e
'   -Ajout de la copie dans le presse papier
'   -Ajout du fusionneur/d�coupeur de fichiers
'
'   v1.3
'   -Support de l'historique + outils
'
'   v1.2
'   -Ajout de la gestion des disques
'
'   v1.1
'   -Ajout de la gestion de la modification des processus en m�moire
'
'   v1.0
'   -Release initiale




' =======================================================
' HISTORY (ENGLISH)
' =======================================================
'
' FINAL RELEASES :
'
'
' BETA RELEASES :
'
'
' ALPHA RELEASES :
'
'
' PRE ALPHA RELEASES :
'
'   v1.7
'   -Added ISO Creator Tool
'   -Added a sanitization tool for disks/files
'
'   v1.6
'   -Various Optimizations of the software (starting, opening of the files...)
'   -Added the management of end of file
'   -Added console
'   -Added new options
'   -Fixed lots of bugs
'   -Added Disassembler Tool and French/English complete tanslations
'   -Added selection and support of physical disks
'   -Fixed display bugs
'   -Optimized loading of File Explorer
'   -Added a bookmark manager
'   -Added export to the clipboard and a file for the export of complete files
'   -Added tree structures of process rather than simple lists
'   -Added "Language Editor Tool" project
'   -Added support of multi-language for Hex Editor VB (except for console)
'   -Added complete french translation
'
'   v1.5
'   -Added file search
'   -Added options
'   -Added icons
'   -Fixed many bugs
'   -Added support of CDFS and UDF formats
'   -Added support of more than 900TB files
'
'   v1.4
'   -Added options
'   -Added search in memory
'   -Added advanced conversion
'   -Added copy to clipboard
'   -Added 'file fusion'
'
'   v1.3
'   -Added history + tools
'
'   v1.2
'   -Added disk gestion
'
'   v1.1
'   -Added support of process virtual memory edition
'
'   v1.0
'   -Initial release







' =======================================================
' AUTEUR & REMERCIEMENTS (FRANCAIS)
' =======================================================
'
' Code enti�rement r�alis� par Alain Descotes (violent_ken)
'
' Certains morceaux de codes cod�s par d'autres personnes ont �t� r�utilis�s.
' Merci � eux, � savoir (ordre alphab�tique) : Brunews (bnAlloc.dll),
' Galain (aide pour le support des disques), PCPT (AfClsManifset),
' Paul Caton (self subclassing pour les usercontrol), Renfield (aide sur plein de
' choses), ShareVB (source de Disassembler_Dll + GetFileBitmap)
'
' La Dll de d�sassemblage est directement issue du travail de ShareVB.
' CAsmProc a �t� enti�rement cod�e par EBArtSoft.




' =======================================================
' AUTHOR & THANKS (ENGLISH)
' =======================================================
'
' Coded only by Alain Descotes (violent_ken)
'
' Few parts of code have been coded by other people.
' Thanks to them (in alphabetic order) : Brunews (bnAlloc.dll), Galain (help to
' add disk support), PCPT (AfClsManifset), Paul Caton (usercontrol self subclassing)
' Renfield (help for several things) , ShareVB (Disassembler_DLL + GetFileBitmap)
'
' Disassembler DLL was coded by ShareVB.
' CAsmProc was coded by EBArtSoft.






' =======================================================
' Merci de noter les modifications �ventuelles apport�es au code ci dessous, avec notamment
' une description pr�cise des modifications, l'auteur (et le moyen de le contacter) ainsi
' que la date et la version du code.
'
' Please add your modifications here, and do not forget to mention a precise description
' of the modifications, the author (and the way to contact him), the date and the version
' =======================================================
