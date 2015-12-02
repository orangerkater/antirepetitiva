#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Martin Borer · text-it Produktdokumentation GmbH
 Language:       German
 Platform:       Win7
 Script Function:
	Benennt in einem Ordner und in allen Unterordner Dateinamen anhand einer
	Excel-Liste in die Zielsprache um.

#ce ----------------------------------------------------------------------------

#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <array.au3>

Func Starteingaben()

   $sprachgruppenhoehe = ceiling(($AnzahlSprachen[0]-1)/6)*25 + 20
   $guihoehe = 10 + 40 + 10 + $sprachgruppenhoehe + 10 + 20 + 10

   GUICreate("Dateiumbennener Sprachauswahl",350,$guihoehe)
	  GUICtrlCreateLabel("In welchem Ordner sind die umzubenennenden Dateien?", 10, 10)
		 $Eingabe_BaseDir = GUICtrlCreateInput($BaseDir,10,30,300,20)
		 $Open_BaseDir = GUICtrlCreateButton("...", 315, 30, 25, 20)

	  GUICtrlCreateGroup("In welche Sprache sollen die Dateien umbenannt werden?",5,60,300,$sprachgruppenhoehe)
		 Dim $SpracheEingabe[$AnzahlSprachen[0]+1]
		 For $i=1 to $AnzahlSprachen[0]-1
			$PosX = 15+($i-(Ceiling($i/6)-1)*6)*45-45
			$PosY = 80+Ceiling($i/6)*25-25
			$SpracheEingabe[$i+1] = GUICtrlCreateRadio($Dateiname[1][$i+1],$PosX,$PosY,40,20)
		 Next
	  GUICtrlCreateGroup("",-99,-99,1,1) ;close group
	  If $Sprache > 0 Then GUICtrlSetState($SpracheEingabe[$Sprache],$GUI_CHECKED)

	  $OK = GUICtrlCreateButton("OK", 25, $guihoehe - 30, 120, 20)
	  $Cancel = GUICtrlCreateButton("Abbrechen", 25 + 10 + 120, $guihoehe - 30, 120, 20)

   GUISetState()

   While 1
	  $msg = GUIGetMsg()
		 IF $msg = $OK Then ExitLoop
		 IF $msg = $Cancel Then Exit

		 IF $msg = $Open_BaseDir Then
			$BaseDir = FileSelectFolder("In welchem Ordner sind die umzubenennenden Dateien?", GUICtrlRead($Eingabe_BaseDir))
			GUICtrlSetData($Eingabe_BaseDir, $BaseDir)
		 EndIf

	  $BaseDir = GUICtrlRead($Eingabe_BaseDir)

	  For $i=1 to $AnzahlSprachen[0]-1
		IF $msg = $SpracheEingabe[$i+1] Then $Sprache = $i+1
	  Next
   WEnd
EndFunc

;;; Array aus TXT befüllen
 Dim $DAZ ;DateinamenAnzahlZeilen
_FileReadToArray(@ScriptDir & "\Dateinamen.txt", $DAZ)
$AnzahlSprachen = stringsplit($DAZ[1],Chr(9),1)
Dim $Dateiname[$DAZ[0] + 1][$AnzahlSprachen[0] + 1]
For $x = 1 to ($DAZ[0])
   $oneRow = stringsplit($DAZ[$x],Chr(9),1)
   For $y = 1 to ($AnzahlSprachen[0])
	  $Dateiname[$x][$y] = $oneRow[$y]
   Next
Next

$BaseDir = ""
$Sprache = 0

Starteingaben()

While $Sprache = 0 OR $BaseDir = ""
   GUIDelete("Dateiumbennener Sprachauswahl")
   MsgBox($MB_SYSTEMMODAL, "Dateiumbenenner", "Keine Sprache oder keinen Pfad ausgewählt!")
   Starteingaben()
WEnd

$Dateiliste = _FileListToArrayRec($BaseDir, "*.doc*" , $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT , $FLTAR_FULLPATH )

$Sicherheitsfrage = MsgBox($MB_YESNO, "Dateiumbenenner", "Sollen wirklich " & $Dateiliste[0] & " Dateien umbenannt werden?" & @CRLF & @CRLF & "Ordner sind die umzubenennenden Dateien:" & @CRLF & $BaseDir )
	  if $Sicherheitsfrage = 7 Then Exit

$CounterF = 0;
$CounterE = 0;

For $Datei = 1 to $Dateiliste[0]
   For $i = 2 to $DAZ[0]
	  If $Dateiname[$i][1] <> "" And StringInStr($Dateiliste[$Datei],$Dateiname[$i][1]) > 1 Then
		 If $Dateiname[$i][$sprache] = "" Then
			$NeuerName = StringReplace($Dateiliste[$Datei],$Dateiname[$i][1],$Dateiname[$i][1] & "__DILFEHLER")
			$CounterF = $CounterF + 1
		 Else
			$NeuerName = StringReplace($Dateiliste[$Datei],$Dateiname[$i][1],$Dateiname[$i][$sprache])
			$CounterE = $CounterE + 1
		 EndIf

		 FileMove($Dateiliste[$Datei],$NeuerName)
		 ExitLoop

	  EndIf
   Next
Next

MsgBox($MB_SYSTEMMODAL, "Dateiumbenenner", "OK. Habe fertig." & @CRLF & @CRLF & $CounterF & " Dateien wurden NICHT umbenannt." & @CRLF & $CounterE & " Dateien wurden umbenannt.")
