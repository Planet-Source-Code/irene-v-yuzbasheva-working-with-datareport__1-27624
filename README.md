<div align="center">

## Working with DataReport


</div>

### Description

Code below demonstrates how to work with DataReport. You will need:

1. DataEnvironment

2. DataEnvironment Command that receives

parameters.

3. DataReport with;

a. Detail section named "Section1"; and,

b. rptlabel named "lblCopy"

4. Form with a button. Code below should be put on click event.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Irene V\. Yuzbasheva](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irene-v-yuzbasheva.md)
**Level**          |Advanced
**User Rating**    |2.9 (38 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/irene-v-yuzbasheva-working-with-datareport__1-27624/archive/master.zip)





### Source Code

```
With UILetterDataReport 'DataReport Name
  .DataMember = "cmdClaimsLetter" 'Name of the Command
Set .DataSource = UIDataEnvironment 'Name of the DataEnvironment
  .Caption = "Letter for: SSN= " & strSSN ' Changing caption properties of the report
 With .Sections("Section1").Controls  'Name of the detail section of the report.
   .Item("lblCopy").Caption = "COPY" 'Set caption of the field
 End With
.Refresh  'Refresh Report
End With
```

