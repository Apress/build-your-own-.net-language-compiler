qbVariableTest

This form and application tests the Quick Basic variable object


C H A N G E   R E C O R D ---------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     ---------------------------------------------
06 26 03   Nilges         Version 1

07 25 03   Nilges         1.  Bug: use of unshared test object 
                                  
10 08 03   Nilges         Bug: progress bar too long and not concealed                                   
                                  
02 08 04   Nilges         1.  Use windowsUtilities.updateStatusListBox                                   
						  2.  Add tool tips and a close box
                                  
02 09 04   Nilges         1.  Use disposeInspect method in place of
                              dispose
                                  
02 10 04   Nilges         1.  Test button should be available when
                              test object is not available
                                  
02 10 04   Nilges         1.  In the test, create 0..32 random test
                              objects
                              
                          2.  Log all MsgBox messages
                          
                          3.  Standardized button dimensions    
                          
                          4.  Added button to change value
                          
                          5.  De-enable controls on Leave and not on
                              text changed event
                                  
02 11 04   Nilges         1.  Extended the About information
                          2.  Improved status Zoom
                                  
03 04 04   Nilges         1.  Added centerToScreen to form Load

05 02 04   Nilges         2.  Replaced netValueEntry by modifyVariable
                
						  

I S S U E S -----------------------------------------------------------
  DATE       POSTER       DESCRIPTION AND RESOLUTION
--------   ----------     ---------------------------------------------
02 09 04   Nilges         When I click Describe, the message claims 
                          that the qbVariableType is inside Form1:
                          this doesn't make sense
                          
02 09 04   Nilges         Creation of random variables does not work
                          
02 10 04   Nilges         NetValueEntry form has not been completely
                          tested: may be further problems
                          
05 06 04   Nilges         1.  Issues with the modifyVariable form
                              1.1 It is cloned in the qbGUI project
                              1.2 It hasn't been thoroughly tested
