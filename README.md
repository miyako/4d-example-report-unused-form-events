![version](https://img.shields.io/badge/version-20%2B-E23089)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-example-report-unused-form-events)](LICENSE)

# 4d-example-report-unused-form-events

[dependencies](https://github.com/miyako/4d-example-report-unused-form-events/blob/main/report-unused-form-events/Project/Sources/dependencies.json)

* [`miyako/Audit`](https://github.com/miyako/Audit)
* [`miyako/xlsx-editor`](https://github.com/miyako/xlsx-editor)

## usage

```4d
var $forms : cs.Audit.Forms
$forms:=cs.Audit.Forms.new()

var $XLSX : cs.XLSX.XLSX
$XLSX:=cs.XLSX.XLSX.new()

var $templateFile : 4D.File
$templateFile:=File("/RESOURCES/événements inutiles.xlsx")

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

$sheet1.A1:={value: "Table"; bold: True; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}
$sheet1.B1:={value: "Form"; bold: True; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}
$sheet1.C1:={value: "Event"; bold: True; size: 14; fill: "FF000077"; stroke: "FFFFFFFF"}

$sheet2.A1:={value: "Table"; bold: True; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}
$sheet2.B1:={value: "Form"; bold: True; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}
$sheet2.C1:={value: "Event"; bold: True; size: 14; fill: "FF007700"; stroke: "FFFFFFFF"}

var $pTable : Pointer
var $tableName; $formName; $cellName : Text
var $tableNumber; $i1; $i2 : Integer
var $formNames; $events; $eventsInMethod : Collection
var $form; $event : Object

$i1:=1
$i2:=1

For ($tableNumber; 0; Last table number)
	ARRAY TEXT($formNamesArray; 0)
	Case of 
		: ($tableNumber=0)
			FORM GET NAMES($formNamesArray)
			$tableName:=""
		: (Not(Is table number valid($tableNumber)))
			continue
		Else 
			$pTable:=Table($tableNumber)
			FORM GET NAMES($pTable->; $formNamesArray)
			$tableName:=Table name($tableNumber)
	End case 
	
	$formNames:=[]
	
	ARRAY TO COLLECTION($formNames; $formNamesArray)
	
	For each ($formName; $formNames)
		$form:=$forms.formDefinition($formName; $tableNumber)
		$events:=$forms.formEvents($form)
		$eventsInMethod:=$forms.formMethodEvents($form)
		For each ($event; $events)
			If ($eventsInMethod.query("constant == :1 and token == :2"; $event.constant; $event.token).length=0)
				$i1+=1
				$cellName:="A"+String($i1)
				$sheet1[$cellName]:=$tableName
				$cellName:="B"+String($i1)
				$sheet1[$cellName]:=$formName
				$cellName:="C"+String($i1)
				$sheet1[$cellName]:=$event.constant
			End if 
		End for each 
		
		For each ($event; $eventsInMethod)
			If ($events.query("constant == :1 and token == :2"; $event.constant; $event.token).length=0)
				$i2+=1
				$cellName:="A"+String($i2)
				$sheet2[$cellName]:=$tableName
				$cellName:="B"+String($i2)
				$sheet2[$cellName]:=$formName
				$cellName:="C"+String($i2)
				$sheet2[$cellName]:=$event.constant
			End if 
		End for each 
	End for each 
	
End for 

var $outputFile : 4D.File
$outputFile:=Folder(fk desktop folder).file("événements inutiles.xlsx")

$XLSX.update({file: $templateFile; values: $values; output: $outputFile})

OPEN URL($outputFile.platformPath)
```

## result

<img width="591" height="708" alt="" src="https://github.com/user-attachments/assets/33388963-8a53-4e1c-aa5f-745eca109b5b" />

<img width="591" height="708" alt="" src="https://github.com/user-attachments/assets/daf9d1d2-9f3a-4b05-99db-0edc5ae06f60" />
