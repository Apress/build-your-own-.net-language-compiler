qbGUI

This is an example Graphical User Interface for the Quick Basic engine which 
foregrounds a glass-box view of the inner workings of the engine.


C H A N G E   R E C O R D -----------------------------------------------------
  DATE     PROGRAMMER     D E S C R I P T I O N
--------   ----------     -----------------------------------------------------
12 14 03   Nilges         Bug: in testing for a rerun, examined the output
						  screen.  

12 14 03   Nilges         Bug: options shouldn't be initialized by getting qbe
                          values.  

12 21 03   Nilges         1. Added code to highlight source in the basic view 

                          2. The status box and progress bar should appear in the
                             basic view 
                             
                          3. Show interpreter events in the status box
                          
                          4. Label the status box: provide zoom capability  

01 04 04   Nilges         1. Added option to suppress the Stop display 

01 06 04   Nilges         1. Bug: constant evaluation option was not saved 

01 09 04   Nilges         1. Scroll the screen to the current entry

01 12 04   Nilges         1. Remove dispose in options form: causes loop

02 17 04   Nilges         1. Add tool tips

                          2. Add close box
                             
                          3. Don't log error when it is in the trial compile, or,
                             modify it to show it's nonfatal   
                             
                          4. The use of the load menu item should toggle a return
                             to the source code screen   
                             
                          5. Exchanging txtScreen's Tag values causes errors: use
                             two item collection to contain source and output
                             
                          6. Added chkCEtestEventLog      
                          
02 24 04   Nilges         1. Bug: did not save source code properly: fixed     

                          2. Bug: clicking replay at end of replay storage
                             caused a crash
                             2.1 Wrapped replay_ call in a null error handler
                             2.2 See also issue with this date                       
                          
02 27 04   Nilges         1. Bug: initial display of large screen made the 
                             status area too small
                             1.1 Changed the showMore method to adjust clientSize
                                 before and after the status control is built:
                                 prior to building the status control, the screen
                                 is deepened to a proportion of available space:
                                 after, it is again adjusted to wrap around the
                                 actual size of the status control.     
                          
02 27 04   Nilges         1. Added resize to fit screen space to more display

02 28 04   Nilges         1. Resize enhancements
							 1.1 Form too narrow
						     1.2 Adjust status list box height
						  2. Add timing to Run and evaluation
						  3. Add test case for variable declaration	 
						  4. Bug: extra newlines in print statements
						     4.1 Changed displayOutput        
                                        
03 02 04   Nilges         1. Adjustments to resize based on flaws in 120 DPI view

03 06 04   Nilges         1. Need the resize object
                          2. Stop button failed to appear (no error message) in
                             120 DPI view
                          3. Abend occured when using run on the menu
                          4. End tags in XML parse not indented
                          5. Got stack underflow in XML parse
                          6. XML parse appeared to make system hang in a thread
                          7. Remove old updateStatusListBox
                          8. Added ParseTrace and ParseOutline options

03 14 04   Nilges         1. Bug: stack underflow in parse trace
                             1.1 Fixed: compiler had mismatched stack entries
                             
03 15 04   Nilges         1. Added a resizer cache           
                          2. Fine-tuned size of controls
                          3. Bug: running initially in the Less display caused
                             an abend (object was nothing)
                             3.1 Problem was that the CTLstatusDisplay is 
                                 needed to create the parse stack
                             3.2 Fixed by adding check for CTLstatusDisplay
						  4. Bug: run form cleanup went past end of Controls
						     collection
						     4.1 Fixed
						  5. When displaying the stop button adjust height of
						     main form to available space
						  6. Save source code in Registry: restore from Registry 
						  
04 08 04   Nilges         1. Bug: display of array did not condense defaults
                             1.1 Fixed the problem by fixing qbVariable
                          3. By default, don't display array values    
                          3. Added QBGUI_EXTENSIONS, and option to display
                             arrays in full  
                          4. Added tooltips to option form
                          5. Added stack pop check option       
                          6. Added QBGUI_POPCHECK
                          7. Made source text box an RTF text box: bold face
                             compiled material, leaving it bold face until
                             it is altered (and, needs to be recompiled)     
						  8. Degenerate op removal option not stored in Registry
						     8.1 Fixed
						     
04 26 04   Nilges         1. Suppressed file menu items in the output mode						                                                   	  

I S S U E S -------------------------------------------------------------------
  DATE     POSTER    FIXED BY  FIX DATE D E S C R I P T I O N            
-------- ---------- ---------- -------- ---------------------------------------   
11 22 03 Nilges                         codeChange(index) overload, and map
										of code described in documentation,
										is not implemented.
										
11 22 03 Nilges                         Don't see modifications to code made
                                        by the compiler.
                                        
11 22 03 Nilges                         See warning re StatusDisplay.
                                        RPNListBox 
                                        
11 23 03 Nilges                         Have added ScanListBox as a stopgap
                                        to be able to highlight tokens 
                                        during parse
                                        
11 24 03 Nilges                         The runInterface, the compileInterface,
                                        and the assembleInterface all scan
                                        and compile UNNECESSARILY: add code
                                        to check to see if the SourceCode
                                        has been changed.                                                                                                                        
                                        
11 24 03 Nilges                         Boxed error messages that contain long
									    words or long file identifiers display
									    incorrectly
									    
11 28 03 Nilges                         The status display returns a number of
										possibly overbroad properties including
										handles to the group box and the RPN
										display as a stopgap: consider refining
										the object model instead to avoid
										pathological changes to the status display.									                                                                                                                           
									    
11 28 03 Nilges                         For large programs the parse display will
										run off the side of the list box
										
11 28 03 Nilges     Nilges     12 14 03 Failed to rerun a compiled program		

12 16 03 Nilges                         1. Need to test the source trace 	
                                        2. Doesn't appear to scroll to caret when
                                           displaying output
                                        3. Added eventLogFormat form
                                        4. In interpretation don't display the scanner
                                           or the parser, so as to have space for the
                                           RPN, stack and storage displays
                                           
01 04 04 Nilges                         The use of the Stop button causes a crash
                                        (button disposed) upon return to the main form.
                                        Have made its use an option. 
                                           
01 04 04 Nilges                         Stop improvements 
                                        1. Immediate response      																                                                                                                                           
										2. Panic button  
										                                         
01 08 04 Nilges                         When an interpreter error occurs the Run still
                                        reports success

01 28 04                                Failed when run with screen whose layout was revised
                                        per Apress specifications for bitmaps. Make work
					                    for this screen	

02 23 04                                Didn't get a progress bar the second time the test
                                        was run using the Test button

02 24 04                                See Change Record (2) for 2-24: should probably
                                        find out reason for error rather than just wrapping
                                        replay statement in null error handling 

02 25 04                                Had to press Stop button a couple of times to stop
                                        running program

02 25 04                                Ordinary message (use of stop button causes abend)
                                        displayed in a manner by the EXE which makes message
                                        appear to be worse than it is 

02 27 04                                When you request the stop button the message 
                                        "not enough screen space" appears 

02 28 04                                Loading files when output screen is being displayed
                                        causes strange results
                                        
02 29 04                                Resize on maximize: this entails more thorough testing
                                        of egnResize and activation of code in qbGUI
                                        archive for 2-29-2004.
                                        
                                        Note that the resizer is needed because repeatedly
                                        toggling from Less to More view causes drift in the
                                        status display.      
                                        
03 10 04                                1. Ran hello world and then loaded nFactorial: did not
                                           compile nFactorial, printed output in source
                                           window                                              
                                           
                                        2. XML output is wrong: for the HelloWorld program it
                                           starts with two repetitions of the print terminal   
                                        
03 11 04                                1. Current version neither saves nor restores the
                                           source code                                              
                                                                                   
03 12 04                                1. Got the following messages when attempting a
                                           forced compile after a Run
                                           
***************************************
* 5 Specified argument was out of the *
* range of valid values.              *
* Parameter name: '75' is not a valid *
* value for 'index'.                  *
***************************************

*****************************************
* 6 Arithmetic operation resulted in an *
* overflow.                             *
*****************************************        

03 12 04                                1. Should have option to suppress ANY parse trace:
                                           RESOLVED                                                                         

03 12 04                                1. Run in same session fails to compile

03 14 04                                1. Cannot load and run new program, have to reload
                                           compiler
                                           
03 15 04                                1. Have developed own resizer, should use
                                           egnResize   

                                        2. See showMore_ensureStatusVisibility: this
                                           fixes a problem such that when the stop button
                                           is visible in the "Less" view, the labels and 
                                           list boxes of the status display are not visible
                                           in the More view.
                                           
                                           Also, use of showmore always turns on command
                                           replay.
                                           
                                           Ideally the status display should respond to 
                                           a make-visible event by making controls visible.                                           
                                           
03 21 04                                1. Previous source code appears in output window   

04 27 04                                1. Clone test menu item produces a fail
                                        2. Allow right click of storage elements to show
                                           their value (including arrays)
                                           
05 06 04                                1. The form modifyVariable.VB is duplicated across
                                           this project and qbVariableTest
                                           
05 10 04                                1. Bug: parsing of parse display failed when the
                                           current token was a delimiter                                                                                      



                                        
                                        	
