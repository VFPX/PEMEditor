* Find the top-most parent object for the specified object. The complication is
* that the top level container is a Formset in the Form and Class Designers,
* and even worse, one level down from the Formset is a Form for non-form
* classes in the Class Designer. So, we need to be careful about finding the
* top-most parent object.
 
* Unfortunately, there's no way to distinguish between a FormSet and a Form
* open in the Form Designer. In both cases, the FormSet is a baseclass. So,
* this code just returns a reference to the form.
 
lparameters toObject, ;
     tlClass
local llClass, ;
     loObject1, ;
     loObject2
 
* Determine if we're in the Class or Form Designer.
 
llClass = iif(pcount() = 1, upper(wontop()) = 'CLASS DESIGNER', tlClass)
store toObject to loObject1, loObject2
 
* Go up the containership hierarchy until we hit a form.
 
do while type('loObject1.Parent.Name') = 'C' and ;
     not upper(loObject1.BaseClass) == 'FORM' and ;
     not upper(loObject1.BaseClass) == 'TOOLBAR'
     
     loObject2 = loObject1
     loObject1 = loObject1.Parent
enddo while type('loObject1.Parent.Name') = 'C' ...
do case
 
* If the form's parent is a non-baseclass FormSet, we're modifying a FormSet in
* the Class Designer, return a reference to the formset.
 
     case upper(loObject2.BaseClass) == 'FORMSET'

     case upper(loObject1.Parent.BaseClass) == 'FORMSET' and ;
          not empty(loObject1.Parent.ClassLibrary)
          loObject2 = loObject1.Parent
 
*** JRN 2009-09-26 : Courtesy of Burkhart Stiller 
* for a SCX based formset, the form does not have 'BufferMode'
     case (not llClass) and upper(loObject1.Parent.BaseClass) == 'FORMSET' 	;
          and Not PemStatus(loObject1, 'BufferMode', 5)
          loObject2 = loObject1.Parent
 
* If the form is a baseclass Form and we're in the Class Designer, we're
* modifying a non-form class, so already have the correct object reference.
 
     case llClass and upper(loObject1.Class) == 'FORM'
 
* For all other cases (a form class open in the Class Designer or a form open
* in the Form Designer), return a reference to the form.
 
     otherwise
          loObject2 = loObject1
EndCase

return loObject2
 