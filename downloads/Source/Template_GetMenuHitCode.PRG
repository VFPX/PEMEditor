lparameters toParameter
local lnSelect, ;
     lcCode, ;
     llReturn
try
     lnSelect = select()
     select 0
     use (_foxcode) again shared order 1
     if seek('M' + padr(upper(toParameter.MenuItem), len(ABBREV)))
          lcCode = DATA
     endif seek('M' ...
     use
     select (lnSelect)
     if not empty(lcCode)
          llReturn = execscript(lcCode, toParameter)
          if llReturn
              toParameter.ValueType = 'V'
          endif llReturn
     endif not empty(lcCode)
catch
endtry
return llReturn
