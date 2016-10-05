creditEvaluation

This form and this application demonstrate a practical application of the
Quick Basic engine to business rules in the form of a credit evaluation
calculator.


C H A N G E   R E C O R D -----------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     -----------------------------------------------
01 05 04   Nilges         Version 1.0  
                                                                                                                                                                                                  
01 09 04   Nilges         1.  Set the text box to wait message during
                              evaluation
                              
                          2.  Explain contradictory rules
                          
                          3.  Get contradictory rules when rerunning
                          
                          4.  Make sure defaultRule is last
                          
                          5.  Complete evaluation options
                                                                                                                                                                                                  
03 17 04   Nilges         1.  Add tooltips and Close button
                              
                          2.  Resize to fit the screen
                          
                          3.  Added Imports statement for the
                              utilities libraries

I S S U E S -------------------------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION AND RESOLUTION
--------   ----------     -----------------------------------------------
01 08 04   Nilges         sortRules facility removed and archived, 
                          produces duplicate rules
                          
01 08 04   Nilges         Consider optimizing by substitution of 
                          values for variable names, and using the
                          constant eval feature of qbe. Note that qbe
                          constant eval has 2 issues at this time:
                          the value cache is updated wrong, and jumps
                          aren't optimized in for and in do 
                                                                               
01 08 04   Nilges         Use of undefined names will cause a crash      
                                                                                                     
01 08 04   Nilges         Don't duplicate default policies         
                                                                                                                        
01 08 04   Nilges         Hasn't been thoroughly tested: needs a stress
                          test  
                                                                                                                                                                            
01 08 04   Nilges         Runs too slowly for a practical solution 
                          1. With or without optimization and with integer 
                             overflow checks takes 11 secs on my machine to 
                             evaluate the default case
                                                                                                                                                                             
01 08 04   Nilges         Test reuse of object       
                                                                                                                                                                             
01 09 04   Nilges         Doesn't seem to reuse the quick basic engine
                          as expected
                                                                                                                                                                             
01 09 04   Nilges         Edit rule screen doesn't show rule

01 09 04   Nilges         A crash occured after multiple evaluations at
                          the point where the evaluation object was
                          evaluated: it was Nothing
                                 
                                