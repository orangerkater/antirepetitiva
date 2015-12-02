#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Martin Borer Â· text-it Produktdokumentation GmbH
 Language:       German
 Platform:       Win7
 Script Function:
	FÃ¼r einen Ordner/ein Laufwerk (zB USB-Stick) ein Inhaltsverzeichnis erstellen,
	als Textdatei (.log), Worddatei(.doc) und PDF(.pdf) abspeichern.

	Eingaben: Aufzulistendes Verzeichnis, Ausgabeverzeichnis und Dateiname.

#ce ----------------------------------------------------------------------------

#include <Constants.au3>
#include <Word.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <array.au3>
#include <Date.au3>

Func Starteingaben()
   GUICreate("Ordnerinhalt auflisten und als .doc & .pdf speichern",350,220)
	  GUICtrlCreateLabel("Welcher Pfad (Ordner/Laufwerk) soll aufgelistet werden?", 10, 10)
		 $Eingabe_AuflistPfad = GUICtrlCreateInput($AuflistPfad,10,30,300,20)
		 $Open_AuflistPfad = GUICtrlCreateButton("...", 315, 30, 25, 20)

	  GUICtrlCreateLabel("Wo sollen die Listen gespeichert werden?", 10, 60)
		 $Eingabe_ZielPfad = GUICtrlCreateInput($ZielPfad,10,80,300,20)
		 $Open_ZielPfad = GUICtrlCreateButton("...", 315, 80, 25, 20)

	  GUICtrlCreateLabel("Auftragsnummer?", 10, 110)
		 $Eingabe_Auftragsnummer = GUICtrlCreateInput($Auftragsnummer,10,130,130,20)

	  GUICtrlCreateLabel("Sticknummer?", 150, 110)
		 $Eingabe_Sticknummer = GUICtrlCreateInput($Sticknummer,150,130,130,20)

	  $OK = GUICtrlCreateButton("OK", 10, 180, 120, 20)
	  $Cancel = GUICtrlCreateButton("Abbrechen", 140, 180, 120, 20)

   GUISetState()

   While 1
	  $msg = GUIGetMsg()
		 IF $msg = $OK Then ExitLoop
		 IF $msg = $Cancel Then Exit

		 IF $msg = $Open_AuflistPfad Then
			$AuflistPfad = FileSelectFolder("Welcher Pfad (Ordner/Laufwerk) soll aufgelistet werden?", GUICtrlRead($Eingabe_AuflistPfad))
			GUICtrlSetData($Eingabe_AuflistPfad, $AuflistPfad)
		 EndIf
		 IF $msg = $Open_ZielPfad Then
			$ZielPfad = FileSelectFolder("Wo sollen die Listen gespeichert werden", GUICtrlRead($Eingabe_ZielPfad))
			GUICtrlSetData($Eingabe_ZielPfad, $ZielPfad)
		 EndIf

	  $AuflistPfad = GUICtrlRead($Eingabe_AuflistPfad)
	  $ZielPfad = GUICtrlRead($Eingabe_ZielPfad)
	  $Auftragsnummer = GUICtrlRead($Eingabe_Auftragsnummer)
	  $Sticknummer = GUICtrlRead($Eingabe_Sticknummer)
   WEnd
EndFunc

;; Startvariablen
$Sticknummer = "1"
$Auftragsnummer = "9xxx"
$AuflistPfad = "E:\"
$ZielPfad = "C:\Users\TechRed\Downloads\temp"

Starteingaben()

$ZielName = $Auftragsnummer & "_" & StringReplace(_NowCalcDate(),"/","-") & "_Ordnerinhalt_USB-Stick " & $Sticknummer

   $Ziel = $ZielPfad & "\" & $ZielName & ".log"
   $ausgabe = FileOpen($Ziel, 2)
	  FileWriteLine($ausgabe, "Ordnerinhalt erstellt am:")
	  FileWriteLine($ausgabe, _Now())
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "Pfad bzw. Ordner:")
	  FileWriteLine($ausgabe, $AuflistPfad)
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "--------------------------------------------------------------------------------------------")
	  FileWriteLine($ausgabe, "")
   FileClose($ausgabe)
   RunWait(@ComSpec & ' /c ' & 'tree "' & $AuflistPfad & '" >> "' & $Ziel & '"', @WindowsDir, @SW_HIDE)
   $ausgabe = FileOpen($Ziel, 1)
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "--------------------------------------------------------------------------------------------")
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "Auflistung der Dateien")
	  FileWriteLine($ausgabe, "")
   FileClose($ausgabe)
   RunWait(@ComSpec & ' /c ' & 'dir /s/o/c/a:-d "' & $AuflistPfad & '" >> "' & $Ziel & '"', @WindowsDir, @SW_HIDE)
   $ausgabe = FileOpen($Ziel, 1)
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "--------------------------------------------------------------------------------------------")
	  FileWriteLine($ausgabe, "")
	  FileWriteLine($ausgabe, "End of File")
   FileClose($ausgabe)

   Global $oWord = _Word_Create()
	  $oWord.Run("DosDatei_öffnen", $ZielName & ".log", $ZielPfad)
	  $oWord.Run("Ordnerinhalt_als_doc")
	  $oWord.Run("AusgabeAlsDocUndPdf_ASpeichern", $ZielName, $ZielPfad)
