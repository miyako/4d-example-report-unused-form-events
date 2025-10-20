//%attributes = {"invisible":true}
var $forms : cs:C1710.Audit.Forms
$forms:=cs:C1710.Audit.Forms.new()

var $XLSX : cs:C1710.XLSX.XLSX
$XLSX:=cs:C1710.XLSX.XLSX.new()

var $templateFile : 4D:C1709.File
$templateFile:=File:C1566("/RESOURCES/événements inutiles.xlsx")

var $values : Object
$values:={}

var $sheet1Title; $sheet2Title : Text
$sheet1Title:="not used in method"
$sheet2Title:="not activated form"
var $sheet1; $sheet2 : Object
$values[$sheet1Title]:={}
$values[$sheet2Title]:={}
$sheet1:=$values[$sheet1Title]
$sheet2:=$values[$sheet2Title]

$sheet1.A1:={value: "Table"; bold: True:C214; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}
$sheet1.B1:={value: "Form"; bold: True:C214; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}
$sheet1.C1:={value: "Event"; bold: True:C214; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}

$sheet2.A1:={value: "Table"; bold: True:C214; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}
$sheet2.B1:={value: "Form"; bold: True:C214; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}
$sheet2.C1:={value: "Event"; bold: True:C214; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}

var $pTable : Pointer
var $tableName; $formName; $cellName : Text
var $tableNumber; $i1; $i2 : Integer
var $formNames; $events; $eventsInMethod : Collection
var $form; $event : Object

$i1:=1
$i2:=1

For ($tableNumber; 0; Last table number:C254)
	ARRAY TEXT:C222($formNamesArray; 0)
	Case of 
		: ($tableNumber=0)
			FORM GET NAMES:C1167($formNamesArray)
			$tableName:=""
		: (Not:C34(Is table number valid:C999($tableNumber)))
			continue
		Else 
			$pTable:=Table:C252($tableNumber)
			FORM GET NAMES:C1167($pTable->; $formNamesArray)
			$tableName:=Table name:C256($tableNumber)
	End case 
	
	$formNames:=[]
	
	ARRAY TO COLLECTION:C1563($formNames; $formNamesArray)
	
	For each ($formName; $formNames)
		$form:=$forms.formDefinition($formName; $tableNumber)
		$events:=$forms.formEvents($form)
		$eventsInMethod:=$forms.formMethodEvents($form)
		For each ($event; $events)
			If ($eventsInMethod.query("constant == :1 and token == :2"; $event.constant; $event.token).length=0)
				$i1+=1
				$cellName:="A"+String:C10($i1)
				$sheet1[$cellName]:=$tableName
				$cellName:="B"+String:C10($i1)
				$sheet1[$cellName]:=$formName
				$cellName:="C"+String:C10($i1)
				$sheet1[$cellName]:=$event.constant
			End if 
		End for each 
		
		For each ($event; $eventsInMethod)
			If ($events.query("constant == :1 and token == :2"; $event.constant; $event.token).length=0)
				$i2+=1
				$cellName:="A"+String:C10($i2)
				$sheet2[$cellName]:=$tableName
				$cellName:="B"+String:C10($i2)
				$sheet2[$cellName]:=$formName
				$cellName:="C"+String:C10($i2)
				$sheet2[$cellName]:=$event.constant
			End if 
		End for each 
	End for each 
	
End for 

var $outputFile : 4D:C1709.File
$outputFile:=Folder:C1567(fk desktop folder:K87:19).file("événements inutiles.xlsx")

$XLSX.update({file: $templateFile; values: $values; output: $outputFile})

OPEN URL:C673($outputFile.platformPath)