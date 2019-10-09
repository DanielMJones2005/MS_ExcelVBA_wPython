# Creating and Executing MS Excel VBA with Python
## Case example of XML Import

Summary: 
This code will create a new MS Excel file, create a new MS Excel VBA Module, add desired VBA code, 
execute VBA code, save MS Excel File, close MS Excel File

### Step 1: Macro Settings
- Macro Settings should allow developer VBA project object model before running

### Step 2: Define VBA Code
```
vba_active_wb = '    ActiveWorkbook.XmlImport URL:= _'
file_path = 'C:\folder\sub-folder\'
file_name = 'xmlFile.xml'
vba_xml_param =f'        , ImportMap:=Nothing, Overwrite:=True, Destination:=Range("$A$A")'
vba_end = 'End Sub"

xml_code = f'''Sub XMLImport()
{vba_active_wb}
        "{file_path}\{file_name}" _
{vba_end}'''

print(xml_code)
```

```
Sub XMLImport()
    ActiveWorkbook.XmLImport URL:= _ 
        "C:\folder\sub-folder\xmlFile.xml" _
        , ImportMap:=Nothing, Overwrite:=True, Destination:=Range("$A$A")
End Sub
```

```
alert_off = "Application.DisplayAlerts = False"

vba_alert_off = f'''Sub AlertOff()
        {alert_off}
{vba_end}'''

print(alert_off)
```

```
Sub AlertOff()
        Application.DisplayAlerts = False
End Sub
```

```
alert_on = "Application.DisplayAlerts = True"

vba_alert_on = f'''Sub AlertOn()
        {alert_on}
{vba_end}

print(alert_on)
```

```
Sub AlertOff()
        Application.DisplayAlerts = True
End Sub
```

### Step 3: Import | Activate Excel
```
import win32com.client

xl = win32com.cleint.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
ss = xl.Workbooks.Add()
```

### Step 4: Add VBA Modules and VBA Code
```
xlmodule = ss.VBProject.VBComponents.Add(1)

xlmodule.CodeModule.AddFromString(vba_alert_off)
xlmodule.CodeModule.AddFromString(vba_alert_on)
xlmodule.CodeModule.AddFromString(xml_code)
```

### Step 5: Run VBA Code from Python
```
xl.Application.Run( r'AlertOff' )
xl.Application.Run( r'XMLImport' )

ss.SaveAs('/folder/sub-folder/newFileName.xlsx')

xl.Application.Run( r'AlertOn' )

ss.Close()
```
    
