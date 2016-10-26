# Jakarta-CityTouch-kWh

## User Management



## Main 
* FromDate is date indicator (same as display on daily bar chart)
* Copy content from all named workbook: maximum 31 sheets per month

* Ref> http://stackoverflow.com/questions/19351832/copy-from-one-workbook-and-paste-into-another

```
Sub foo()
Dim x As Workbook
Dim y As Workbook

'## Open both workbooks first:
Set x = Workbooks.Open(" path to copying book ")
Set y = Workbooks.Open(" path to destination book ")

'Now, copy what you want from x:
x.Sheets("name of copying sheet").Range("A1").Copy

'Now, paste to y worksheet:
y.Sheets("sheetname").Range("A1").PasteSpecial

'Close x:
x.Close

End Sub
```

### Setup
* Pre-add blank Sheets("Input") manually as a placeholder for daily energy report 
     * :o: ? Wanna make it for all sites or just one site? 
* Pre-add blank Sheets("Temp") manually as temporary process data placeholder

### Process
* Process one file
     * Copy the 1st file Sheets("Sheet1").UsedRange to Workbook("Processor").Sheets("Input")
     * :o: ? How to determine the sequence of the files
* 
* If Gardu value is empty (e.g. Gardu = "-")?
     * Replace "-" with "Default" and consider it as one Gardu "Default"
*
```

```

* Obtain the Unique Gardu values:
     * Copy Sheets("Input").Column(Gardu) to Sheets("Temp").Column(Gardu) with unique value is true
* Set Sheets("Temp").Column(kWh).formula = "Vlookup"
* Find the TotalGarduQty = lastEmptyRow of Sheets("Temp"), 
* Set Sheets
* Clean up: Workbook("Processor").Sheets("Input").UsedRange.delete



## Admin
* Load Func library 

* Sheets("Ref") as the configuration page to hold the Public Variables from the NamedRange and the Public constants 
    * Sheets folder path: text field, filled by operator 
        * Sheets folder path: can be file selection window: Ref> http://stackoverflow.com/questions/10304989/open-windows-explorer-and-select-a-file
    * Site: drop down list, select by operator
    * Sheets("Sheet1").Range("C2").Value is "site name" e.g. "Jakarta Pusat"
    * Sheets("Sheet1").Range("E3").Value is "From date", but uses rightest 10 chars only
    * Sheets("Sheet1").Range("F3").Value is "To date", but uses rightest 10 chars only 
    * Sheets("Sheet1").Range("D3").Value is Query name, should be constant :"%All lums kWh export%"
    * Columns, Module wide constants
        * Energy Consumption kWh	: "A"
        * Serial Number (Luminaire)	: "B"
        * Lamp Burning Hours	: "C"
        * Installation Date (Luminaire)	: "D"
        * Group	: "E"
        * Longitude	: "F"
        * Latitude	: "G"
        * Kecamatan	: "H"
        * Gardu	: "I"
        * IDPEL: "J"


## Pending features
* Additional 2 columns to record the FromDate and ToDate to cross check whether any missing date is there
* Incorrect data checking (:o: whether want to enable it as pre-check?)
* Cross check the Gardu and Group ID/name? 
* Whether the operator can select a group of files and then process it one by one?
 * http://stackoverflow.com/questions/12687536/how-to-get-selected-path-and-name-of-the-file-using-open-file-dialog-control
```
Sub Demo()
    Dim lngCount As Long
    Dim cl As Range

    Set cl = ActiveCell
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            ' Add Hyperlinks
            cl.Worksheet.Hyperlinks.Add _
                Anchor:=cl, Address:=.SelectedItems(lngCount), _
                TextToDisplay:=.SelectedItems(lngCount)
            ' Add file name
            'cl.Offset(0, 1) = _
            '    Mid(.SelectedItems(lngCount), InStrRev(.SelectedItems(lngCount), "\") + 1)
            ' Add file as formula
            cl.Offset(0, 1).FormulaR1C1 = _
                 "=TRIM(RIGHT(SUBSTITUTE(RC[-1],""\"",REPT("" "",99)),99))"


            Set cl = cl.Offset(1, 0)
        Next lngCount
    End With
End Sub
 ```
 

