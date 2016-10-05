integerCalc: Integer calculator (Professional/Enterprise Edition 2003 code)

This form and application supports a simple calculator that adds, subtracts, multiplies,
and divides integer values.  It illustrates simple lexical analysis and parsing, using
informal regular expressions and BNF as a specification language.

Note that the compile-time symbol TRACE may be set to True in the project properties:
when this symbol is True, extra controls will allow the user to trace scanning, parsing
and interpretation.

Note also that when TRACE is False, the project Reference to utilities should be removed.

C H A N G E   R E C O R D ---------------------------------------------
  DATE       PROGRAMMER   DESCRIPTION OF CHANGE
--------     ----------   ---------------------------------------------
01 24 03     Nilges       Version 1.0

01 26 03     Nilges       Version 2.0

                          1.  Vers 1 copies to folder "Final vers-
                              without trace"
                              
                          2.  Added trace feature    
                          
                          3.  Added saving and restoring of expression
                              to/from Registry

07 22 03     Nilges       Version 2.1

                          1.  Bug: when replay is clicked prior to any
                              evaluation, application crashes
                              1.1 Check COLmacro for Nothing

07 22 03     Nilges       Updated references

01 11 04     Nilges       Version 3

                          1. Replay cleanup
                             1.1 Always create the COLmacro collection
                             1.2 Reset the COLmacro collection for each
                                 evaluate
                             1.3 Be sure that when step is executed and
                                 the macro pointer is 0 or beyond the end
                                 of COLmacro, the list boxes
                                 are cleared, and step returns to the
                                 beginning  
                          2. Add minimize, maximize and close buttons          

02 27 04     Nilges       Set ClientHeight and ClientWidth, in place of
                          Height and Width
                              