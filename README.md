# Jakarta-CityTouch-kWh

## User Management



## Main 

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

## Admin
* Load Func library 

* Sheets("Ref") as the configuration page to hold the Public Variables from the NamedRange and the Public constants 
> Sheets folder path: text field, filled by operator
> Site: drop down list, select by operator
> Sheets("Sheet1").Range("C2").Value is "site name" e.g. "Jakarta Pusat"
> Sheets("Sheet1").Range("E3").Value is "From date", but uses rightest 10 chars only
> Sheets("Sheet1").Range("F3").Value is "To date", but uses rightest 10 chars only 
> Sheets("Sheet1").Range("D3").Value is Query name, should be constant :"%All lums kWh export%"
> Columns, Module wide constants
>* Energy Consumption kWh	: "A"
>* Serial Number (Luminaire)	: "B"
>* Lamp Burning Hours	: "C"
>* Installation Date (Luminaire)	: "D"
>* Group	: "E"
>* Longitude	: "F"
>* Latitude	: "G"
>* Kecamatan	: "H"
>* Gardu	: "I"
>* IDPEL: "J"

> 


* If Gardu value is empty (e.g. Gardu = "-")?
* Cross check the Gardu and Group ID/name? 

