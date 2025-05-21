# MS Excel things

Showcase of some of the MS Excel things I've previously done.

## Contents

1. [Attribute Collator](#attribute-collator)
2. [G-code Snake Pattern Generator](#g-code-snake-pattern-generator)
3. [Organization Hierarchy Formatter](#organization-hierarchy-formatter)

---

# Attribute Collator

![Screenshot of first sheet in "Attribute Collator.xlsx"](https://github.com/jsjs2401/ms-excel-things/blob/main/images/Attribute%20Collator.png)

Collates attributes from input data where each entry corresponds to one attribute for a specific item, with the output presenting all available attributes in one row for each item.

Also removes empty and duplicate attributes.

### Equivalent code in Python:

```python
input = [["water", "wet"], ["water", "slippery"], ["fire", "hot"], ["water", "refreshing"], ["water", "refreshing"],
         ["fire", "panas"], ["wood", "hard"], ["fire", "atsui"], ["metal", "harder"], ["wood", ""], ["metal", "heavy"]]


def attrCollator(input):
    output = dict()
    for x in input:
        if x[1] == "":
            continue
        if x[0] in output.keys():
            if x[1] not in output[x[0]]:
                output[x[0]].append(x[1])
        else:
            output[x[0]] = [x[1]]
    return output
```

---

# G-code Snake Pattern Generator

![Screenshot of "G-code snake pattern generator.xlsm" alongside the output g-code file](https://github.com/jsjs2401/ms-excel-things/blob/main/images/G-code%20snake%20pattern%20generator.png)

Generates a snaking pattern for characterizing gel behaviour in a pneumatic syringe 3D printer (specifically the CellInk BioX, as that was what I worked with). Supports multimaterial printing. The gel extrusion rate is set at a constant "E1" in the code, but can be modified to suit filament extrusion printers by scaling it proportionally to the print speed.

I made this mostly because it's hard to get slicers programs to behave fully consistently when you change the dimensions of the 3D model (sometimes it does diagonal movements instead), and sometimes I just want one layer of snaking pattern. It probably could have been more neatly coded in Python, but this was also for other people in my lab to use, since everyone has MS Excel installed, but not everyone has Python installed.

It's meant to print snaking patterns like shown immediately below, although multimaterial cuboid structures can also be printed with the right settings. Multimaterial printing in this file is determined from the set percentage ("Percent Material 2"), using VBA's Rnd function, although it could be easily modified to print the different materials from a 2D lookup array if you have your own function to generate one. The pattern shown in the cube print was generated from a sum-of-sines algorithm.

![Example image of the snaking pattern that can be printed](https://github.com/jsjs2401/ms-excel-things/blob/main/images/G-code%20snake%20pattern%201.png)
![Example multimaterial prints of cubic structures using different settings in the pattern generator](https://github.com/jsjs2401/ms-excel-things/blob/main/images/G-code%20snake%20pattern%202.png)

<details>
         <summary><h3>VBA Code, if you don't want to download the Excel macro file</h3></summary>

```vbnet
Private baseLen, baseWid, lineSpace, printHeight, layerHeight As Double
Private printSpeed1, printSpeed2, printhead1, printhead2, gCode As String
Private numLayers, numLinesTotal, curPrinthead, curMat1Line, curMat2Line As Integer
Private percentMat2 As Single

Private startPosX, startPosY, posX, posY, posZ As Double

Sub snakeCube2Pattern()

'Initializing parameters from sheet into code
With Sheet3
    baseLen = .Range("B1").Value
    baseWid = .Range("B2").Value
    lineSpace = .Range("B3").Value
    printHeight = .Range("B4").Value
    layerHeight = .Range("B5").Value
    printSpeed1 = " F" & .Range("B6").Value * 60
    printSpeed2 = " F" & .Range("B7").Value * 60
    numLinesTotal = WorksheetFunction.Floor_Precise(baseWid / lineSpace, 1) + 1
    printhead1 = "T" & .Range("B8").Value - 1
    printhead2 = "T" & .Range("B9").Value - 1
    percentMat2 = .Range("B10").Value
End With

'Checks that all sheet parameters are filled
If baseLen = 0 Or baseWid = 0 Or lineSpace = 0 Or printHeight = 0 Or layerHeight = 0 Or printSpeed1 = "" Or printSpeed2 = "" Or printhead1 = "" Or printhead2 = "" Then
    Message = MsgBox("Make sure all values are filled up or non-zero.", vbOKOnly)
    Exit Sub
End If

'Calculating and setting other variables from parameters
numLayers = WorksheetFunction.Ceiling_Precise(printHeight / layerHeight, 1)
startPosX = baseWid / 2
startPosY = baseLen / 2
gCode = "M83 ;Relative extrusion mode" & vbNewLine & _
        "G21 ;Metric values" & vbNewLine & _
        "G90 ;Absolute positioning" & vbNewLine & _
        "M107 ;Start with the fan off" & vbNewLine & _
        "G28 ;Home the printer" & vbNewLine & _
        "G92 E0 ;Zero the extruder" & vbNewLine

'Generate full g-code according to parameters
Dim i As Integer
For i = 1 To numLayers
    posZ = Round((i - 1) * layerHeight, 2)
    nextStr = "G1 Z" & posZ & " F2400"
    gCode = gCode & vbNewLine & nextStr
    gCode = gCode & vbNewLine & snakeLayer(i)
Next i

'Lifts up from the final point by 5mm
nextStr = "G1 Z" & layerHeight * numLayers + 5 & " F2400"
gCode = gCode & vbNewLine & vbNewLine & nextStr

'Temporary for testing, will change to output to file when done
'Sheet3.Range("A15").Value = gCode

'Saves file in selected location
SaveFileAs

End Sub

Sub SaveFileAs()

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fileName As Variant
Dim overwriteCheck As Integer

SelectSaveAs:
fileName = Application.GetSaveAsFilename(fileFilter:="g-code (*.gcode), *.gcode")
If fileName <> False Then
    If fso.fileExists(fileName) Then overwriteCheck = MsgBox("File already exists. Overwrite file?", vbYesNo)
    If overwriteCheck = vbNo Then GoTo SelectSaveAs
    
    Dim txtfile As Object
    Set txtfile = fso.CreateTextFile(fileName, True)
    txtfile.WriteLine (gCode)
    txtfile.Close
    Set txtfile = Nothing
End If
Set fso = Nothing

End Sub

Function snakeLayer(i) As String 'Generates gcode for each layer

posX = startPosX
posY = altNum(i - 1) * startPosY 'Sets the starting X and Y positions for this layer
posZ = Round((i - 1) * layerHeight, 2)

Dim snakeLayerMat1, snakeLayerMat2 As String
Dim j, prevMat As Integer

If i <= 1 Or i >= numLayers Then 'For first or last layer, only print with Material 1
    snakeLayer = printhead1 & vbNewLine & "G1 X" & startPosX & " Y" & altNum(i + 1) * startPosY & " Z" & posZ & printSpeed1 'Starting position
    For j = 1 To numLinesTotal 'Prints snake pattern, going up and down Y axis, starting at +ve X and +/-ve Y
        posY = startPosY * altNum(j) * altNum(i + 1)
        nextStr = "G1 X" & posX & " Y" & posY & printSpeed1 & " E1"
        snakeLayer = snakeLayer & vbNewLine & nextStr
        
        If j < numLinesTotal Then 'Adds extrusion movement across X if not last line
            posX = Round(startPosX - j * lineSpace, 2)
            nextStr = "G1 X" & posX & " Y" & posY & printSpeed1 & " E1"
            snakeLayer = snakeLayer & vbNewLine & nextStr
        End If
        
    Next j

Else 'For non-edge layers, randomize the material according to the desired percentage:
    snakeLayerMat1 = printhead1 & vbNewLine & "G1 X" & startPosX & " Y" & altNum(i + 1) * startPosY & " Z" & posZ & printSpeed1
    posY = startPosY * altNum(i)
    nextStr = "G1 X" & posX & " Y" & posY & printSpeed1 & " E1"
    snakeLayerMat1 = snakeLayerMat1 & vbNewLine & nextStr 'Print first line as Material 1, starting at +ve X and +/-ve Y
    
    prevMat = 1
    snakeLayerMat2 = printhead2
    
    For j = 2 To numLinesTotal - 1 'Prints snake pattern, going up and down Y axis, starting at +ve X and +/-ve Y
        Randomize
        If Rnd > percentMat2 Then
            If j <= numLinesTotal Then 'Adds extrusion movement across X if not last line
                posX = Round(startPosX - (j - 1) * lineSpace, 2)
                nextStr = "G1 X" & posX & " Y" & posY & printSpeed1
                If prevMat = 1 Then nextStr = nextStr & " E1"
                snakeLayerMat1 = snakeLayerMat1 & vbNewLine & nextStr
            End If
            prevMat = 1
            
            posY = startPosY * altNum(i + 1) * altNum(j)
            nextStr = "G1 X" & posX & " Y" & posY & printSpeed1 & " E1"
            snakeLayerMat1 = snakeLayerMat1 & vbNewLine & nextStr
        
        Else
            If j <= numLinesTotal Then 'Adds extrusion movement across X if not last line
                posX = Round(startPosX - (j - 1) * lineSpace, 2)
                nextStr = "G1 X" & posX & " Y" & posY & printSpeed2
                If prevMat = 2 Then nextStr = nextStr & " E1"
                snakeLayerMat2 = snakeLayerMat2 & vbNewLine & nextStr
            End If
            prevMat = 2
            
            posY = startPosY * altNum(i + 1) * altNum(j)
            nextStr = "G1 X" & posX & " Y" & posY & printSpeed2 & " E1"
            snakeLayerMat2 = snakeLayerMat2 & vbNewLine & nextStr
            
        End If

    Next j
    
    posX = Round(startPosX - (numLinesTotal - 1) * lineSpace, 2)
    nextStr = "G1 X" & posX & " Y" & posY & printSpeed1
    If prevMat = 1 Then nextStr = nextStr & " E1"
    snakeLayerMat1 = snakeLayerMat1 & vbNewLine & nextStr
    posY = startPosY * altNum(i + 1) * altNum(numLinesTotal)
    nextStr = "G1 X" & posX & " Y" & posY & printSpeed1 & " E1"
    snakeLayerMat1 = snakeLayerMat1 & vbNewLine & nextStr 'Prints last line as Material 1
    
    snakeLayer = snakeLayer & vbNewLine & snakeLayerMat1 & vbNewLine & snakeLayerMat2
    
End If

End Function

Function altNum(var) As Integer 'Returns 1 if odd numbers, -1 if even numbers

altNum = (-1) ^ (var Mod 2)

End Function
```

</details>

---

# Organization Hierarchy Formatter

![Screenshot of first sheet in "Organization Hierarchy Formatter.xlsx"](https://github.com/jsjs2401/ms-excel-things/blob/main/images/Organization%20Hierarchy%20Formatter.png)

Converts relational organizational data into a chart form.

Includes some error-checking such as:

- Checking for duplicate organizations.
- Checking for potentially missing relational links.
- Some capacity to identify corrupted input data.

### Equivalent code in Python:

```python
from collections import deque

input = [["HQ", ""], ["Office 1", "HQ"], ["Office 2", "HQ"], ["Office 3", "HQ"], ["Office 4", "HQ"], ["Office 5", "HQ"],
         ["Office 6", "HQ"], ["Suboffice 1-1", "Office 1"], ["Suboffice 1-2", "Office 1"], ["Suboffice 1-3", "Office 1"],
         ["Suboffice 2-1", "Office 2"], ["Suboffice 3-1", "Office 3"], ["Suboffice 3-2", "Office 3"],
         ["Suboffice 4-1", "Office 4"], ["Suboffice 5-1", "Office 5"], ["Suboffice 6-1", "Office 6"]]


def orgHierarchyFormatter(input):
    organizations = dict()
    output = dict()
    for x in input:
        if x[0] in organizations.keys():
            print(f"Warning: Duplicated organization {x}")
        else:
            organizations[x[0]] = x[1]
    for org in organizations.keys():
        if org not in organizations.values() and org not in output.keys():
            output[org] = deque([org])
    for x in output.keys():
        while output[x][0] != "":
            try:
                output[x].appendleft(organizations[output[x][0]])
            except KeyError:
                print(f"Warning: Missing relational link for {output[x][0]}")
                break
    return output
```
