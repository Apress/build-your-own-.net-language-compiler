quickBasicEngine

This class does all scanning, parsing and interpretation for my version
of Quick Basic. It may be "dropped in" to a .Net application and it 
will provide the ability to evaluate immediate Basic expressions as well
as compile and run Basic programs.


C H A N G E   R E C O R D ---------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     ---------------------------------------------
08 17 03   Nilges         Version 1.0 (supersedes older nonobject
                          version written for VB.Net 2002)
                          
10 12 03   Nilges         1.  Bug: missing Return(True) in mulFactorRHS
                          
                          2.  Suppress Polish object comments in XML
                          
11 22 03   Nilges         1.  Bug: compared unprefixed token name to
                              prefixed token name
                              
                          2.  Various bugs in using indexes to tokens
                              and Polish operations were fixed
                              
                          3.  Added the codeChangeEvent
                          
                          4.  Parse event must report terminal status
                          
                          5.  Added token start index to parse event    
                          
                          6.  Fixed issue "bombs out when closed on probable
                              inspection error"
                              
                          7.  Added PolishCollection    
                          
                          8.  AssemblyRemovesCode did not default correctly
                          
                          9.  Bug: in clearing storage, use of resetVariable
                              reset the type of the variable: replaced
                              resetVariable with clearVariable
                              
                          10. Bug: default step value set incorrectly as a
                              fromString on a null qbVariable: use 
                              mkVariableFromValue    
                          
12 13 03   Nilges         1.  Bug: the Scanner property was documented but
                              not implemented.
                              
12 14 03   Nilges         1.  Bug: the duplicate implementation did not use
                              the operand    
                              
12 14 03   Nilges         1.  Changed object2XML    

12 16 03   Nilges         1.  Bug: multiply was not supported

12 18 03   Nilges         1.  Evaluate bug fixes
                              1.1 Clone needed to test for an existing scanner
                              1.2 Clone needed to return object
                              1.3 CompareTo needed to test for an existing 
                                  scanner
                                  
                          2.  The inspect method may now omit the report parameter.   
                          
                          3.  Added EventLogging, and EventLog properties     
                          
                          4.  Added eventLogFormat method     

12 20 03   Nilges         1.  Added InspectCompilerObjects

12 21 03   Nilges         1.  Efficiency issues

							  1.1 Removed the conversion of the qbVariable to
							      its value in the interpretation of pushLiteral
							      to save 1..2 seconds in 5! benchmark
							      
						  2.  Multithread issues
						  
						      2.1 Make fully threadable
						      
						      2.2 Added properties and methods for threads
						      
						      2.3 Added Tag property
						      
						      2.4 Added qbe to each event		
						 
01 02 04   Nilges         Added limited MSIL generation 
						 
01 05 04   Nilges         1.  Added overload to evaluate to obtain .Net values 
						  2.  Added evaluationValue
						 
01 06 04   Nilges         1.  Bug: invalid values combined in a constant eval 
						 
01 08 04   Nilges         1.  Bug: Incorrect test applied to loop event raiser's result
                              value 
						 
01 08 04   Nilges         1.  Bug: Incorrect handling of Boolean operators needed to 
                              push -1 for True and 0 for False for Jumpz
						 
01 17 04   Nilges         1.  Bug: failed on variant creation after updates to
                              qbVariant
                          2.  Error messages need to be formatted for readability
						 
01 22 04   Nilges         1.  Bug: tag set did not send its operand to the dispatcher,
                              fixed
						 
01 24 04   Nilges         1.  Removed the collectionClearHandler because it creates
                              excessive output
                              
                          2.  Made event logging an option of the test method because
                              it creates excessive output
                              
                          3.  Ordered Select cases in dispatch_ by approximate 
                              frequency        
                              
01 24 04   Nilges         1.  Bug: failed to handle ampersand suffix properly
                              1.1 Fixed: the key requirement: the parser must test the
                                  type suffix to make sure it is physically adjacent
                                  to the identifier.        
                              
01 29 04   Nilges         1.  Support for As clause
                          2.  Don't add keywords to the variable table
                          3.  Bug: failed to handle numeric suffix correctly
                              3.1 Fixed: as above the parse must test the suffix
                                  to be sure it is adjacent to the number
                                  
03 02 04   Nilges         1.  Added parseStartEvent and parseFailEvent   
   
                          2.  Event documentation does not identify the sender
                              parameters.  
                              
                          3.  Problems in compiler__functionCall_
                              3.1 Found out of order Return statement
                              3.2 Documentation header wasn't standard
                              
                          4.  Bug found and fixed: exit parse return identified
                              the exit as an unconditionalStatement                                  
                              
                          5.  Bug found and fixed: relOp parse return identified
                              the relOp as an optionStatement                                  
                              
                          6.  Bug found and fixed: powOp parse return identified
                              the powOp as an optionStatement                                  
                              
                          7.  Bug found and fixed: moduleDefinition parse return 
                              identified the moduleDefinition as an 
                              implicitAssignment        
                              
03 12 04   Nilges         1.  Widened zoom boxes for 120 DPI                                                        
                              
03 14 04   Nilges         1.  Bug: concat factor parse created errors in parse
							  display       

03 14 04   Nilges         1. Bug: stack underflow in parse trace
                             1.1 Fixed: compiler had mismatched stack entries

03 17 04   Nilges         1. Failed to compile Print 0 And 1

                          2. Found erroneous gc names in parse events: cf. followup
                             issue dated 3-21    
                             
                          3. Bug: an extra empty line at the end of the data type
                             test program caused the parser to crap out    
                             
                          4. Got the following sequence of messages when pressed 
                             Run twice for a null program
                             
                             4.1 5 Error at 3/20/2004 1:43:40 PM: Invalid token 
                                 index 0 is less than one : Location: qbScanner0001 
                                 3/20/2004 1:43:39 PM.QBToken get: Returning Nothing:
                                 6 Arithmetic operation resulted in an overflow. 

                          5. Optimization (constant folding or assembly) causes
                             data type test program to fail   

                          6. Optimization failed to show any time savings in the
                             test method 

                          7. Added chapter 9 test cases DONE

                          8. Add progress reporting for XML generation DONE

                          9. Optimization bugs
                          
                             9.1 Constants not folded in complexExpression.BAS
                                 FIXED
								 9.1.1 Added code to push immediate command
								       result on the stack
							 9.2 Duplicate key FIXED
								 9.1.1 Added Try with no catch
								 9.1.2 Code incorrectly placed value in index
								 9.1.3 Constant expression table format was 
								       screwy in XML
                             9.2 Degenerate operations not removed in
                                 complexExpression.BAS		
                             9.3 Verify reuse of constant expressions DONE
                             9.4 Verify that all constant expressions are cached    
                                 DONE
                             
                          10. Use booReadable:=True in serializing collections for
                              object2String DONE   
                             
                          11. Test engine should inherit optimization settings DONE  
                          
                          12. Bug: subroutine call to compiler__parseFail_ needed to
                              be function call
                              12.1 Fixed
                              12.2 Cf. issue 3-21
                              
						  13. Bug: compilation of AndAlso/OrElse uses JumpZ and
						      JumpNZ which clear the stack
						      13.1 Stack duplication added	          
						      
						  14. Bug: Tab() function implemented incorrectly
						      14.1 Fixed                        

04 06 04                  1.  Documentation corrections

                          2.  Corrected parameter name of objOpIndex in the code
                              remove event: it is now intOpIndex		
						  
04 08 04   Nilges         1. Bug: don't test for an as clause inside the lvalue
                             handler unless this is a def mode call
                             1.1 Fixed	
						  2. Serialization of large arrays is time-consuming
						     2.1 Test for array and serialize only its type
						         by default
						     2.2 Expose storage2String option to include array
						         values        
						  3. Compilation bombed out at attempt to compile w(0)
						     as array reference
						     3.1 Fixed: bug in qbVariable: need to keep tag
						  4. Get Runtime error at opcode 7: Exception thrown in 
						     interpreter: 5 Error at 4/10/2004 1:25:00 PM: 
						     Compiler error: array subscript stack frame is invalid      
						     4.1 Index is -1
						     4.2 compiler__expressionList_object2IntegerSub_ called
						         with object type, should be called with qbVariable
						         type
						     4.3 As of 1:43 PM 4-10, stack frame looks ok, but,
						         there is still a pause after IP 6, and an error
						         still occurs.       
						         4.3.1 Clone of array seems to be the time sink.
						         
						               Basically, arrays shouldn't be pushed on 
						               the stack for modification by pushIndirect
						               (and the documentation of pushIndirect needs
						               to make it clear that a copy is made).
						               
						               Added functionality to pushIndirect to push
						               reference to UDT or array and not its cloned
						               value.
						               
						               In the interpreter, added code to skip
						               dispose when the value was pushed by reference
						         4.3.2 Changed stack2String to show references   
                             4.4 Test case was in error, need to have two subscripts
                                 4.4.1 Add check for subscript count	
                             4.5 Combine stack frame build with check, allow A to be
                                 part of multiple entries
                             4.6 Efficiency concern: 25! takes 109 secs    
                                 4.6.1 Successful test takes 7 seconds
                                 4.6.2 Added PopCheck property    	
									   4.6.2.1 As result:
									           4.6.2.1.1 25! took 88 seconds without
									                     the pop check: note that there
									                     was an unexpected speed up later
									                     on in the loop: cache?
									           4.6.2.1.2 Test took 6 seconds
                             4.7 Status as of 4-17-2004 8:50 AM
                                 4.7.1 Add test messages re pop check
                                 4.7.2 Test efficiency
									   4.7.2.1 5 test cases take 9, 10, 14, 22, 22 secs
									   4.7.2.2 Unless PopCheck is in effect, bypass the
									           stack frame
                                 4.7.2 Resume array and maze testing 	
                                       4.7.2.1 Added interpretTraceEventFastStack and
                                               storage2StringFastStack
                             4.8 Status as of 4-23-2004 7:24 AM
								 4.8.0 Added ExtensionAvailable and PopCheckAvailable properties
                                 4.8.1 Efficiency issues
									   4.8.2 Option to NOT generate opRem/opNop
											 4.8.2.1 Added GenerateNOPs
									   4.8.3 PopCheck suppression does not enhance 
									         interpreter speed
							     4.8.2 ForTest abends: invalid collection index		   	
							 4.9 Status as of 4-24-2003 5:02 PM
							     4.9.1 It appears pop isn't working and does not decrement the
							           stack top in non-POPCHECK mode
							           4.9.1.1 Fixed
							 4.10 Status as of 4-25-2004 12:32 AM
								  4.10.0 Stack underflow (for test pop off)
										 4.10.0.1 Fixed
							      4.10.1 Need to stress test changed qbVariable
							             4.10.1.1 Done
							 4.11 Status as of 4-25-2004 6:08 PM
							      4.11.1 Test fails (on last case) when constant folding and
							             degenerate op removal are in effect
							             4.11.1.1 Also, fails to visually compile this example when
							                      options are defaulted
							             4.11.1.2 When compiling "Dim strText", the dim remains
							                      highlighted as an error and compilation fails         
							                      after source program body
							      4.11.2 Color coding is ineffective when terminals adjacent and
							             one char long       
							             4.11.2.1 FIXED
							      4.11.2 Transfer corrections from opPopToArrayElement to
							             pushArrayElement
							      4.11.3 Ability to get array dump on right click     
							 4.12 Status as of 4-26-2004 8:05 AM
								  4.12.0 Need to color comments specially
								         4.12.0.1 Done
								  4.12.1 Need to show assembler effect on RPN
								         4.12.1.1 Done
							      4.12.2 Developing code needed to clear source code including Tag
							             reset
							             4.12.2.1 In this, don't expose ability to clear or load
							                      code when in output mode  
							                      4.12.2.1 Done		
							 4.13 Bug: location and length of comment gs incorrect
							      4.13.1 Fixed		
							 4.14 Status as of 4-26-2004
							      4.14.1 Fails when extra newline occurs in source: fix 
							             4.14.1.1 Fixed: note revision to syntax of openCode:
							                      sourceProgramBody after logicalNewline is
							                      now optional
							      4.14.2 Still need to fix test failure under optimization     					                      					         							     
							             4.14.2.1 DONE: needed to allow qbVariable to assign
							                      strings TRUE or FALSE to Boolean variables
							 4.15 Bug: Variable and VariableCount not implemented
							      4.15.1 Added ability to retrieve by name to Variable
							 4.16 Placed object sequence interlocking inside the name create
							 4.17 Made the Variable property read-write             

05 09 04   Nilges         1. Added the interpretClsEvent 
                          2. Added parser code to compile CLS							                
                              				      

I S S U E S -------------------------------------------------------------------
  DATE     POSTER    FIXED BY  FIX DATE D E S C R I P T I O N            
-------- ---------- ---------- -------- ---------------------------------------   
11 22 03 Nilges     Nilges     11 23 03 Bombs out when closed on probable
                                        inspection error
                                                                                           
11 22 03 Nilges                         Problems in parse display

                                        1. Sequence number is not incremented
                                           in Dewey Decimal number
                                           
                                        2. Newline appears literally and needs                                                       
                                           to be converted to graphical format
                                           
                                        Runs slow and jerky 
                                                                                                                                 
12 14 03 Nilges                         1. May not process null lines properly

12 16 03 Nilges                         The way the compiler sets the TokenStartIndex
                                        property of the Polish token is a "misnomer"
                                        since the compiler in gencode generally sets the
                                        index to one past the end of the tokens,
                                        that correspond to the grammar category for
                                        which code is being generated.
                                        
12 20 03 Nilges                         Issues in event log formatting

                                        1.  Newlines and nongraphics should be translated
                                            to display format
                                            
                                        2.  Strange lines when length exceeds 40
                                        
12 20 03 Nilges                         Efficiency issues

										1.  Runs slowly when simple screen is on display:
										    the compiled exec takes FOURTEEN seconds 
										    to calculate 5 factorial                                
                                        
01 06 04 Nilges                         Constant folding issues

			                            1.  Constants inside the print list are folded 
			                                even when the option is off.                               

			                            2.  Cf. compiler__statementBody_if and do, because
			                                in both I had to turn constant folding OFF.
			                                
			                                This is because at this writing no change is
			                                made to the code generation when constant
			                                folding is ON.
			                                
			                                Change code to not loop in the DO or just
			                                generate the right code when constant
			                                folding is in effect.                               
                                        
01 08 04 Nilges                         Undefined variables cannot appear in Print statements,
                                        because the Print command expects a scalar, and cannot
                                        convert the default value of Variant,Null
                                        
01 08 04 Nilges                         Cf. assignsReservedWord.BAS: is able to assign a value
                                        to True.
                                        
01 08 04 Nilges                         Cf. 01082004.BAS: causes an error (add to collection failure?)
                                        and then an overflow results
                                        
01 08 04 Nilges                         Cf. businessRules.BAS: when run with constant evaluation, probably
                                        will fail when it doesn't find the second string "DECLINE" & " "
                                        but then adds it with this as the key
                                        
01 08 04 Nilges                         Cf. structuredIf.BAS: failed to compile when the End statement
                                        was on the same line as the Print statement with this
                                        message:

***************************************************
* 5 Error at 1/9/2004 12:50:55 AM: intIndex -1 is *
* not valid Location: utilities.item Returning a  *
* null item                                       *
*************************************************** 

                                        This message was followed by "arith operation resulted in an
                                        overflow"

01 09 04 Nilges                         XML generation takes a long time (consider changing
                                        QBgui?

01 22 04 Nilges                         1.  Creation of event report takes a long time and has no
                                            status report

                                        2.  Closing may take some time when large collections exist.
                                            Consider supporting ways to avoid erasing large
                                            collections on an itemized basis

02 11 04 Nilges                         1.  When tried to convert 1+1 to MSIL got "Not able to
                                            translate this code to Intermediate Langauge: 13
                                            Cast from the string "Enter N" to type "Double"
                                            is not valid.

02 22 04 Nilges                         1.  Test status: runs cases 1..3: for first program in
                                            case 4: need to set strResult to intercepted output

02 24 04                                Dim i As Integer failed to compile: leaves object in state
                                        that fails inspection

02 24 04                                Doesn't appear as if the type prefixes are implemented,
                                        as types
                                        
02 24 04                                Add comments to test output

03 05 04                                Formal parameter not implemented yet

03 12 04                                When no code loaded get "object reference not set too and
                                        "instance..." message
                                        
03 18 04                                Removed cache initialization from updateConstantCache_,
                                        archived version with this initialization. The cache
                                        should be initialized OK by the compiler reset proc.
                                        
03 19 04                                Loaded datatypeTest.BAS over pre-existing code:
					                    system didn't recompile as it is supposed to                                        
									    
03 21 04                                Still finding errors in the standard, required parser
					pattern: develop inspector tool									    
																			    
03 21 04                                Was not able to use run as a function: did not compile

03 23 04                                Need to systematically match BNF in the book with BNF
                                        in comments and as supported including a match of the
                                        functions implemented and documented in the BNF
                                        
03 23 04                                Check actual behavior of the run function against the
                                        documentation in Appendix A      
                                        
03 23 04                                Need to verify correct implementation of pushArrayElement,
                                        and make sure it corresponds to Chapter 8 text   

03 26 04                                For 123.4/(5.1 + 3)*46.004 the result 700.851103287157

04 02 04                                Run function always fails: found no implementation in the
                                        interpreter
                                        
04 06 04                                Test conformity to discussion of DegenerateOpRemoval property.

04 06 04                                Documentation of interpret method says "by default, interpreter
                                        will inspect temporary stack variables when they are
                                        disposed". Then, nothing: why is this statement here?

04 06 04                                Appendix B documentation of reset method is incomplete

04 06 04                                StopQBE operation is crude. Can't threads be stopped "OUTSIDE"
                                        the threads?                 
                                        
04 16 04                                It doesn't appear that the QUICKBASICENGINE_EXTENSION
                                        removal of the string length limit is implemented  

04 17 04                                Repeated tests take more and more time

04 26 04                                Following message occurs in histogram parse

*************************************
* 5 Error at 4/26/2004 10:45:44 AM: *
* intIndex -1 is not valid          *
*                                   *
* Location: utilities.item          *
*                                   *
* Returning a null item             *
*************************************                                                                                                                                             
                                                                                                                