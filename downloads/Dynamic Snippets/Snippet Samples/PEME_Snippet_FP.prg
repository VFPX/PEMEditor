* Tore Bleken 11-1-2011

****************************************************************
****************************************************************
******
****** First Section: Compile-time constants -- modify as needed
******
****************************************************************
****************************************************************

* snippet-keyword: case insensitive, NOT in quotes
#Define Snippet_Keyword      FP

* delimiter between parameters, IN QUOTES; if empty, only one parameter
#Define Delimiter_Char       '='

* minimum number of parameters to be accepted
#Define Min_Parameters      1

* maximum number of parameters to be accepted
#Define Max_Parameters      2




****************************************************************
****************************************************************
******
****** Middle Section: Setup and cleanup code:  DO NOT CHANGE!!!
******
****************************************************************
****************************************************************

Lparameters lcParameters, lcKeyWord

Local loParams As Collection
Local lcParams, lnI

Do Case
            * if no parameters passed, this is a request for Help
      Case Pcount() = 0
            Return PublishHelp()

            * Only process our keyword 
      Case Pcount() = 2 And Not Upper ([Snippet_Keyword]) == lcKeyWord
            Return .F. && not mine!

      Otherwise
            lcParams = _oPEMEditor.ExtractSnippetParameters (lcParameters, Delimiter_Char, [Snippet_Keyword], Min_Parameters, Max_Parameters)
            If 'C' = Vartype (lcParams)
                  Return Process (&lcParams)
            Endif
Endcase


Function CreateHelpResult (lcSyntax, lcSummaryHTML, lcDetailHTML)
      Local loResult As 'Empty'
      loResult = Createobject ('Empty')
      AddProperty (loResult, 'Name', [Snippet_Keyword])
      AddProperty (loResult, 'Syntax', Evl (lcSyntax, ''))
      AddProperty (loResult, 'Summary', Evl (lcSummaryHTML, ''))
      AddProperty (loResult, 'Detail', Evl (lcDetailHTML, ''))
      Return loResult
Endproc




****************************************************************
****************************************************************
******
****** Last Section: Custom code for this snippet
******
****************************************************************
****************************************************************

#Define CR   Chr(13)
#Define LF   Chr(10)
#Define CRLF Chr(13) + Chr(10)
#Define Tab  Chr(9)
#Define CursorSeparator [>]
* Put your code here to process the parameters passed; make sure to set parameters appropriately
* Result is one of following:
*   character string -- string to insert into edit window, replacing the snippet there (usual case, by far)
*                       if the characters '^^^' are found, they indicate where the cursor is to be placed
*   .T.              -- handled, but edit window already updated (advanced use)
*   .F.              -- not handled

Function Process
      Lparameters lcAlias, lcPrefix
      
      Local laFields[1], lcFieldName, lcNewAlias, lcOpenFilePRG, lcResult, lcThisPRG, lcType, lnDecimals
      Local lnFieldCount, lnI, lnWidth
      Local lnSelect, lcTempfile, lcDeleted, lnDeleted, lcSafety, lcTarget
      lcTarget=Getwordnum(lcAlias,2,CursorSeparator)
      lcAlias=Getwordnum(lcAlias,1,CursorSeparator)

      Do Case
            Case Pcount() = 1
                  lcPrefix = ''
            Case Not Empty (lcPrefix)
                  lcPrefix = lcPrefix + '.'
      Endcase

      If Not Used (lcAlias)
            lcThisPRG     = Sys(16)
            lcThisPRG     = Substr (lcThisPRG, 1 + At (' ', lcThisPRG, 2))
            lcOpenFilePRG = [PEME_OpenTable.PRG]
            If File (lcOpenFilePRG)
                  lcNewAlias = PEME_OpenTable (lcAlias) && may return an different alias
                  If 'C' = Vartype (lcNewAlias)
                        lcAlias = lcNewAlias
                        If Pcount() = 1
                             lcPrefix = lcAlias + '.'
                        Endif
                  Endif
                  If Not Used (lcAlias)
                        Return .T. && mine, nothing to change
                  Endif
            Else
                  Messagebox ('[' + lcAlias + '] not found')
                  Return .T. && mine, nothing to change
            Endif
      Endif

   lcDeleted=Set("Deleted")
   lcSafety=Set("Safety")
   Set Deleted Off 
   lcTempfile='PEM_fieldpicker'
   Select (lcAlias)
   lnSelect = Select()
   Copy To (lcTempfile) structure extended
   Select 0
   Use (lcTempfile)
   Delete all
   On Key Label ENTER Keyboard '{ctrl+w}' Plain 
   On Key Label SPACEBAR do deleteshift
   On Key Label CTRL+A do recallall
   On Key Label CTRL+R do reversedelete
   On Key Label A do recallall
   On Key Label R do reversedelete
   Go Top 
   Define Window PEM_fieldpick From 10,10 Size 1,1 Name PEM_fieldpick
   PEM_fieldpick.Top=50
   PEM_fieldpick.Left=50
   PEM_fieldpick.Width=230
   PEM_fieldpick.Height=70 +(Min(Reccount(),15)*18)

   PEM_fieldpick.Caption='PEMEditor field picker'
   Browse In Window PEM_fieldpick Save Name pem_grid Nowait;
      Fields field_name:r,type=field_type + ':' + Transform(field_len)+Iif(field_type='N',','+Transform(field_dec),''):r
   pem_grid.Left=0
   pem_grid.Top=0
   pem_grid.Scrollbars=2
   pem_grid.deletemark=.F.
   pem_grid.Width=PEM_fieldpick.Width
   pem_grid.Height=PEM_fieldpick.Height
   pem_grid.SetAll('dynamicbackcolor','iif(deleted(),Rgb(192,255,192),Rgb(255,255,255))','column')
   pem_grid.SetAll('dynamicforecolor','iif(deleted(),Rgb(0,0,0),Rgb(192,192,192))','column')
   pem_grid.column1.Width=150
   pem_grid.column1.Header1.Caption='Field name'
   pem_grid.column2.Width=50
   pem_grid.column2.Header1.Caption='Type'

   Activate Window PEM_fieldpick 
   Inkey(0.2)
   Browse Last 
   Release Windows PEM_fieldpick
   On Key Label SPACEBAR 
   On Key Label ENTER
   On Key Label CTRL+A
   On Key Label CTRL+R 
   On Key Label A
   On Key Label R 
   If !Empty(lcTarget)
      lcResult    = [Select]
   Else 
      lcResult    = []
   Endif 
   Count For !Deleted() To lnFieldCount
   If lnFieldCount=0
      Delete All
      lnFieldCount=Reccount()
   Endif 
   Count For Deleted() To lnFieldCount
   lnI=1
   Scan For deleted()
            lcResult    = lcResult                               ;
                  + IIf (Empty (lcResult), [], Tab)        ;
                  + lcPrefix + Trim(field_name)                  ;
                  + IIf (lnI < lnFieldCount, [, ;] + CRLF, [])
      lnI = lnI + 1 
   Endscan 
   If !Empty(lcTarget)
      lcResult = lcResult + [ from ] + lcAlias 
      If !Empty(lcPrefix)
         lcResult = lcResult + ' ' + Trim(lcPrefix,1,'.')
      Endif 
      lcResult = lcResult + [ where ^^^ into cursor ] + lcTarget + [ readwrite]
   Endif 
   Use In (lcTempFile)
   Select (lnSelect)
   If lcDeleted='ON'
      Set Deleted On
   Endif 
   Set Safety Off
   Erase (lcTempfile+'.*')
   If lcSafety='ON'
      Set safety On
   Endif 
   Return lcResult
Endproc

****************************************************************

Function recallall
Local lnrecord, lnMarked
lnRecord=Recno()
Count For Deleted() To lnMarked
If lnMarked<=Reccount()/2
   Delete all
Else    
   Recall all
Endif    
Go lnRecord
Endfunc 
****************************************************************
Function reversedelete
Local lnrecord, lnMarked
lnRecord=Recno()
Scan
   If Deleted()
      Recall 
   Else
      Delete
   Endif 
Endscan    
Go lnRecord
Endfunc 

****************************************************************

Function deleteshift
If Deleted()
   Recall
Else
   Delete
Endif 
Endfunc 

****************************************************************

* Publish the help for this snippet: calls function CreateHelpResult with three parameters:
*    Syntax
*    Summary
*    Full description

* Note that all have these may contain HTML tags

Function PublishHelp
      Local lcDetailHTML, lcSummaryHTML, lcSyntax

      lcSyntax = 'Alias[>TargetCursor][=Prefix]'

      Text To lcSummaryHTML Noshow Textmerge
            Creates a list of selected fields from <u>Alias</u> (for SQL-SELECT)<br>
            If TargetCursor is given a complete Select statement is created
      Endtext

      Text To lcDetailHTML Noshow Textmerge
            Inserts the list of the selected fields from <u>Alias</u> (an open cursor or table in DataSession 1), 
            one per line, as used by SQL-SELECT.  Field names look like Alias.Fld1, Alias.Fld2, etc.   <br><br>
            To use a different prefix for the field names, use the equal sign and the prefix (Alias=OtherPrefix) <br><br>
            To not have any prefix for the field names, use just the equal sign (Alias=) <br><br>
            To create a complete SQL Select statement, include the target cursor name (Alias>Target) <br><br>
            If the table is not open, PEME_OpenTable.PRG is called to allow you to open the table; make modifications
            to it as desired to fit your own needs.<br><br>
      Use the Spacebar to select a field<br>
      Use A or Ctrl-A to select or deselect All fields.<br>
      Use R or Ctrl-R to Reverse the current selection.
      Endtext

      Return CreateHelpResult (lcSyntax, lcSummaryHTML, lcDetailHTML)
Endproc
