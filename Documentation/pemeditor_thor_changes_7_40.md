# Changes to PEMEditor (version for Thor)
## general
To allow work with git on bash, all filenames are converted to lower case.
There was a mixup of cases, so on bash it was impossible to numerate files with *?.
## bug fixes
### found a bug in Create LOCALs that blocked the recognition of existing variables.
Variables declared in the following manner where not found:
```
LOCAL;
 lvVar1

LPARAMETERS;
 lvVar1
 
...
```
The beautify class parses the Keywords with `GETWORDNUM(String,Index)`.
This will fail on the code above because the semicolon will be included.
A new method `GetWordNumX()`, sharing the parameters with *GetWordNum* is introduced,
to replace it and parse out the trailing semicolon.
It also fixes the problem that `GETWORDNUM()` w/o 3th paramater does not include CR - 0h0D as delimiter.
The help says so, but it does not.
The problem with CR exists in `GETWORDCOUNT()` too, so a method `GetWordCountX()` is introduced.
Not all occurrences of *GetWordNum* and *GetWordCount* are replaced.   
The failure was a general one.

### found a bug that stopped recognition of *EndText*
Similar to above, the missing of a space or EOL could stop the parsing of some special keywords. 
Those are typical end-of-block keywords without following code on the same line. Like:
```
 ..
ENDTEXT&&Some Comment
```
This is valid code, *GetWordNum* would return *ENDTEXT&&Some*.
The characters for inline comment must be parsed out. Since this may occur in the form
```
LOCAL;&& Some remark
```
as well, the code is added to *GetWordNumX*.

## Improvements
### Create LOCALs
#### LOCAL single on line
Since the original code came with `LOCAL;`, LOCALs inserted should follow this style.
A new option is created.
#### Align AS
Option to align *AS* clauses for newly inserted LOCALs.   
Setting _Maximum column for 'AS'_ to 0 (**option of BeautifyX**) now acts like _no limit_.
#### Beautify on keywords
Keywords inserted with new LOCAL declaration are now beautified.
#### Space after comma
as on other places of BeautifyX, if, or if not a space is inserted befor semicolon, is controled by BeautifyX option 'Add space before commas'.
#### Height of Thor config form
The Create LOCALs settings class forces a resize of Thors setting form
to allow all options be visible.   
#### This changes.
- ThorRepository
   - clscreatelocals OF thorrepository\source\procs\thor_options_createlocals.vcx
- PEMEditor
   - peme_preferences OF pemeditor\downloads\source\peme_preferences.vcx
   - beautify OF pemeditor\downloads\source\peme_beautify.vcx
   - pemeditor\downloads\source\editpropertydialogenglish.h
   - pemeditor\downloads\source\pemeditorversion.h   
   
---------------
### BeautifyX, SELECTs / Updates
#### inverted loop
The regexp driven parser in beautify::BeautifySelectMain was working with a lot of inserting
and playing with offsets to find the next position.
A simple `For lnI =  loMatches.Count To 1 STEP -1` does the same work
without the need of changing offsets.   
The `STUFF` in the same place lend to odd spacings. Replaced. Old lines remain as comment.
#### Indent ON like JOIN
Option to move the *ON* keyword on the same indentation level as *JOIN*.
While bearable in a form:
```
JOIN xys;
 ON lExpr1;
JOIN ddd;
 ON lExp2
```
I found it superfluous in the alternative form:
```
JOIN xys;
JOIN ddd;
 ON lExp2
 ON lExpr1;
```
Now there is an option.
#### Optional Union Indentation
BeautifyX owns an option to indent a *SELECT* following an *UNION*.
Other style would be indentation of all non *UNION* lines.
That is what this option is for.
```
 SELECT;
  Field,;
  Field;
  FROM table;
UNION;
 SELECT;
  Field,;
```
Lines prefacing the *UNION*. are indentet now.   
Use of parameters for the indentation: 
- Indent for SELECT after UNION
   - This option only indents the *SELECT* line itself. Makes not much sense here.
   - Unused / invisible if new option is set
- Indent for SELECT or UNION
   - Reused 'Indent for SELECT / UPDATE'
   - Double function
      - If there is no *UNION* in the SELECT SQL, the indentation of the *SELECT* statement,
   normally leftmost.
      - On combined SELECT SQL with *UNION*,
   it will swap to indentation of the *UNION* statement. Now intended as leftmost.
- Indent SELECT for combined SELECT
   - the reused 'Indent for UNION'
   - The indentation of *SELECT* statement on the combined SELECT SQL
#### Use blocks of AS phrases
This will use an alternative way to align *AS* phrases.
The idea is align only *AS* groups that belong together like fields or tables.
Example:
```
  SELECT;
     Cur1.Field1 AS nF1,;
     1           AS nF2;
     FROM MyFancyTable     AS Cur1;
     INNER JOIN OtherTable AS Cur2;
     ON .F.;
UNION;
  SELECT;
     Cur3.FieldName1 AS nF1,;
     1               AS nF2;
     FROM Table1       AS Cur3;
     INNER JOIN Cursor AS Cur4;
     ON .T.;
     INNER JOIN cxy    AS Cur5;
     ON .T.
```
#### Split into separate lines, options
New options added to prohibit splitting of 
- SELECT items
- FROM tables
- ORDER BY fields
- GROUP BY fields
- other comma seperated fields, if existing   
Those options can be turned of as a group.
#### UPDATE, placement of SET
If set was placed on the wrong position, indentation failed.   
The following form is still parsed wrong:
```
UPDATE;
 Cur1 SET;
     nF2 = xx
```

#### TabIndex, alignment
TabIndex and alignment of options form reworked. Was a bit wobbly while switching.
#### This changes.
- ThorRepository
   - clsbeautifyx OF thorrepository\source\procs\thor_options_beautifyx.vcx
- PEMEditor
   - peme_preferences OF pemeditor\downloads\source\peme_preferences.vcx
   - beautify OF pemeditor\downloads\source\peme_beautify.vcx
   - pemeditor\downloads\source\editpropertydialogenglish.h
#### Todo
- SELECT, UPDATE, REPLACE and the like will not be splited if written as single line.
  - beautify::BeautifySingleLine does not check this.
  - It needs at least one continuation.
- UPDATE, SET
  - See above


## Files changed:
- ThorRepository
   - clscreatelocals OF thorrepository\source\procs\thor_options_createlocals.vcx
   - clsbeautifyx OF thorrepository\source\procs\thor_options_beautifyx.vcx

- PEMEditor
   - pemeditor\downloads\source\editpropertydialogenglish.h
   - peme_preferences OF pemeditor\downloads\source\peme_preferences.vcx
   - beautify OF pemeditor\downloads\source\peme_beautify.vcx
*
   - pemeditor\downloads\source\pemeditorversion.h

Last changed: _2023/02/06_ ![Picture](../docs/pictures/vfpxpoweredby_alternative.gif)