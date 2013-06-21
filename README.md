#Microsoft Access VBA snippets

Small & higly commented Visual Basic subroutines for Microsoft Access databases. Topics cover a wide variety of applications and are created as a learning tool in order to easily add them into your own applications.
All scripts written by [Adam Wilbert](http://adamwilbert.com) and are free to use, dissect, explore and reappropriate however you wish.
***

##backupBackend.vba
Creates a backup of a back-end database. First determines the current path and name of the back-end file, then appends an abbreviated day of the week to the backup file and places a copy in a `\backup\` subfolder.


##hideRibbon.vba
Simple one line code to attach to a button On-Click event in order to hide (or show) the Access Ribbon toolbar. Can also be triggered by a startup form's On-Load event in order to hide the ribbon when the database launches.

