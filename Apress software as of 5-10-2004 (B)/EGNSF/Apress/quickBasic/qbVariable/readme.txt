qbVariable

This class represents the value of a Quick Basic variable.      
For more information, see its documentation in qbArray.vb.


C H A N G E   R E C O R D ---------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     ---------------------------------------------
06 26 03   Nilges         Version 1

07 22 03   Nilges         Version 2

                          1.  Added fromDope and toDope
                          2.  Added offset
                          3.  In fromString, use existing dope when
                              fromString creates an isomorph
                          4.  Added compatible method
                          5.  Added stress test
                          
08 24 03   Nilges         Version 2.1

                          1.  Expanded the fromString syntax to
                              allow internal parentheses not
                              expressed in XML and quoted strings.    
                              
08 30 04   Nilges         Converted to qbVariable from qbArray        

09 07 04   Nilges         Version 3 completed, and archived on Jason
                          Cotton's system at C:\EGN\EGNSF\APRESS\
                          QUICKBASIC\QBVARIABLE\QBVARIABLETEST\
                          ARCHIVE\09072003\VERSION3.
                          
                          There are seven issues associated with this
                          version listed below but it will function
                          as the variable engine                                              

09 07 03   Nilges         Version 3.1                               
                 
                          1.  Added inspection during test
                          
10 03 03   Nilges         Version 4

                          1.  Version 3.1 archived to C:\egnsf\Apress\
                              qbVariable\Archive\10042003\version3dot1\
                              qbVariable.VB
                              
                          2.  Added UDTs      
                              2.1 Procedures changed
                                  2.1.1 clearVariable
                                  2.1.2 clone
                                  2.1.3 compareTo
                                  2.1.4 empiricalDope
                                  2.1.5 fromString      
                                  2.1.6 inspect
                                  2.1.7 isomorphicVariable
                                  2.1.8 mkRandomVariable
                                  2.1.9 object2XML
                                  2.1.10 test
                                  2.1.11 value
                                  2.1.12 valueSet
                                  
                          3.  This version uploaded to Apress FTP on
                              10/6/2003
                              
10 08 03   Nilges         In toString, output strings must be decorated
                          when they contain non-graphic characters    
                          
10 11 03   Nilges         Removed references and need for qbToken                                                           

10 06 03   Nilges         1.  Uploaded to the Apress Web site

10 08 03   Nilges         1.  Added ability to parse object2String's decorated format   

11 29 03   Nilges         1.  Don't return asterisk for a scalar value's default,
							  return the default value	
							  
11 30 03   Nilges         1. Inspection changed to allow null variants

						  2. Bug: variant assignment failed: changed valueSet__
						     to get type and value	

						  3. Bug: in clearing a Variant the type may not be
						     qbVariable	

						  4. Use object2String and not enquote to decorate value in
						     mkVariableFromValue

						  5. Bug: IsANumber returned False for a number Variant	

						  5. Added isScalar method	
							  
12 20 03   Nilges         1. Added booInspect parameter to dispose method
							  
02 08 03   Nilges         1. Adapted to disposeInspect method of qbVariableType   
                      
02 08 03   Nilges         1. Bug: DA found that "Array,Variant,1,2:Integer(1), 
                             String("A")" fails!
                             1.1 FIXED incorrect index usage in reconstruction
                                 of the decorated value in
                                 fromString__parse_quickBasicDecoValue_
                      
02 09 03   Nilges         1. Bug: Boolean value was converted to quoted True/False
							 1.1 Added code to toString__scalarToString_ to
							     convert Booleans	  	
							     
						  2. Bug: shouldn't use _INTsequence separately from its
						     interlocked increment
						     2.1 Now used only inside the increment
						     
						  3. Added error when scan of fromString is not complete
						  
						  4. Not necessary to mkUnusable when a scan fails   	
						  
						  5. Bug: failed to parse "Array,String,1,10:”A”(10)" with
						     Smart Quotes and a repeater
						     5.1 Now supporting Smart Quotes in qbScanner
						     5.2 Repaired repeater     
						  
						  5. Bug: failed to parse examples in book
							 5.1 Examples
							     5.1.1 Array,Integer,1,2,1,2:(1,2),(3,4)
							     5.1.2 Array,Integer,1,2,1,2,1,2:((1,2),
							           (3,4)),((5,6),(7,8))
						     5.2 Repaired parsing of the value     
						  
						  5. Bug: invalidly converts variant's contained type
						     to the narrowest type    
						     5.1 Fixed: variant valueSet_ needed to use the
						         type in the variant and not the narrow type
						         
						  6. Made object IDisposable, added disposeInspect
						  
						  7. Don't inspect on an internal dispose  
						  
						  8. Didn't form default expression for empty array
						     8.1 Added DefaultValue property and variableType2String
						         method      
						         
						  9. Don't use Shared events
						  
						  10. Simplified test method: it must be executed on
						      an allocated instance and it always creates a test
						      object.     		
						      
						  11. Bug. In quickBasicEngine's Nutty Professor
						      interpreter, the .value of a qbVariable was
						      fetched for a variant Byte and the underlying
						      variable object and not its scalar value was
						      returned.
						      
						      Whereas the mission of the variable method is to
						      return the scalar value. Dammit.  
						      
						      11.1 Changed value method: dereferences .objValue
						           if it is a qbVariable  				         
						      
						  12. Bug. The type of the Variant in Variant,Type...
						      failed to correspond to the type of its value as 
						      in Variant,Null:vtByte(25): this should be
						      Variant,Byte:vtByte(25).
						      
						      12.1 Changed valueSet__ method to force the type  
						      12.2 Added inspection rule to make sure types of
						           variants remain consistent				         
						      
						  13. Bug. Test case, where an array should have contained
						      a long and an integer, contained two Longs.
						      
						      13.1 Fixed fromString__parse_quickBasicDecoValue_
						           to correctly convert values
                          						  
02 11 2004   Nilges    1. Bug: when array collection extended, Remove failed:
						  remove past end of collection
						  1.1 Added check
						  
					   2. Bug: scalar not morphed to collection when comma
					      found in data list without explicit type
					      2.1 Added variable/variableType calculus methods
					          2.1.1 qbVariable.attachVariable
					          2.1.1 qbVariableType.attachVariableType	  
					          
					   3. Added a nice cache
					   
					   4. Added Friendly State property		
					   
					   5. Added timing to test method			          
					   
					   6. Added new("") syntax			          
                          						  
02 14 2004   Nilges    1. Bugs in UDT handling 
                          1.1 fromStrings for UDT failed to convert to 
                              and from the palindromic toString
                              1.1.1 In toString, needed to get value of
                                    qbVariable values
						  1.2 Multimember fromstrings with values cause
						      crash (unusable object)
						      1.2.1 String needs to be quoted when
						            the member object value is formed
						            as an inner qbVariable
					   2. Bugs in 1-d array handling
					      2.1 Needed to attach comma-separated members
					          in a loop						            
					   3. Bugs in multiple-dimension array handling
					   4. Bug: allowed comma delimited list for scalar  
					   5. Bug: allows Array,Byte,0,1,0,1:(1,2),("") without
					      complaining about type (causes inspect error)
					   6. Bug: assignment of variant values
					      6.1 Fixed
					   7. Bug: did not create quoted string variants
					      7.1 Needed to enquote internal fromString	
					   8. Bug: 32768 failed to convert
					      8.1 Needed to assign objNew in parser's setValue
					   9. Bug: correct type not created in array
					      9.1 Needed to force type when appending to 
					          collection
					   10. Non-null variable types with errors should be
					       flagged, as errors					          					      					      				            
					   11. Bug: some tests preserved attributes of 
					       previous tests, therefore, recreation of
					       variable created using "32768", using 32767
					       was Long and not Integer as required					          					      					      				            
			               11.1 Added resetVariable to test method               
                          						  
02 22 2004   Nilges    1. Changes in use of qbScanner.findRightParenthesis
                          were due to a change in this method: it now always
                          starts at parenthesis level one
                          						  
04 05 2004   Nilges    1. Documentation corrections

04 08 2004   Nilges    1. Bug: testing of (*) against default was incorrect
                          in toString_
                          
                          1.1 Fixed: in toString_, needed to change the way
                              the booNondefault switch was set
                              
                       2. Bug: from string Array,String,0,1,0,1:("",""),
                          (""),"") was changed to Array,String,0,1,0,1:(),
                          (vtString(""),) 
                          
                          2.1 Fixed
                          
                       3. Don't reset the Tag when the variable is cleared.         

04 23 2004   Nilges    1. Removed test for string overflow: moved to interpreter

04 27 2004   Nilges    1. Added the code that allows the strings True or False
                          to be assigned to Boolean variables.
                          
05 02 2004   Nilges    1. Added optional overload to toMatrix for specifying
                          decoration                                                        
                          
                          
I S S U E S -----------------------------------------------------------
  DATE       POSTER             DESCRIPTION             RESOLUTION
--------  ------------ ----------------------------- ------------------  
07 25 03  Nilges       Hasn't been fully tested

09 07 03  Nilges       1. Does not work for variants 
                          containing arrays, have 
                          removed test 6 temporarily
                          Cf. variant contains 
                          array.txt in archive  
                          
                          Note that when variants
                          containing arrays are
                          supported, inspection 
                          will have to be changed
                          to allow this.
                          
                       2. Add tests of random values
                       
                       3. Inspection of variants, 
                          arrays and UDTs needs to be 
                          more intensive; in particular,
                          apply to them the toString
                          serialization rule   
                          
                       4. Doesn’t apparently handle 
                          repeat count. Cf. Repeat 
                          Count Test.txt in archive 
                          9/7   
                          
                       5. Cf. TEST_CONTAINMENT.TXT 
                          in archive 9-8: containment 
                          not tested	   
                          
                       6. Ability to stress test 
                          would be nice   

                       7. mkRandomVariable will create
                          unsupported variables 
                          including variants containing
                          arrays
                          
                       8. Added isAnUnsignedInteger
                       
09 14 2003   Nilges    1. Added mkVariableFromValue

10 04 2003   Nilges    1. UDTS can, at this writing, never
                          be contained or isomorphic
                          
                       2. toMatrix has not been tested   
                       
                       3. object2XML doesn't give enough
                          information about objValue as a
                          collection
                          
                       4. May be a problem in fromString__parse_
                          setValue_ when object is converted
                          to a string using CStr, may pass
                          unquoted string   
                          
                       5. Decorated notation fails in array specs   
                          
                       6. When the variable is created, and you
                          close immediately with or without
                          dispose, you get the inspection
                          failure (try Array,Variant,1,10:
                          1,"A")   
                          
                       7. Get an inspection failure, for a type/value
                          accepted as ok, for an array with two few
                          entry values
                          
                       8. Use of repeater fails, try
                          Array,String,1,10:"A"(10)
                          
                       9. Array,Integer,1,2,1,2,1,2:((1,2),(3,4)),((5,6),(7,8))
                          did not create three-dimensional array
                          
                      10. 1,2,3 did not parse and did not create an array     
                      
                      11. May need to support indexes inside period-separated
                          member name lists in value and valueSet
                          
                      12. Could not clone a variable
                      
                      13. Could not "create random variables"
                      
11 30 2003   Nilges    1. Wasn't able to rebuild project after making null
						  variable change
						  
12 17 2003   Nilges    1. While testing the compiler, I saw qbVariable
                          toString outputs that looked strange including
                          Variant,Byte: vtDouble(1): why is the concrete
                          Variant type so different from the type in the
                          value??
                          						  
02 08 2004   Nilges    1. Executing the Clear method on an array causes
                          it to fail inspection.						  
                          						  
02 08 2004   Nilges    1. Entered Array,Integer:1,2,3: this was considered
                          invalid. Aren't lower and upper bounds deduced
                          from the array values?						  
                          						  
02 08 2004   Nilges    1. Syntax error was found in UDT,(Member1,Variant,Long),
                          (Member2,Variant,Long),(Member3,Variant,Boolean):
                          Variant,Long:vtLong(0),Variant,Long:vtLong(0),
                          Variant,Boolean:vtBoolean("False")						  
                          						  
02 10 2004   Nilges    1. Cf. change record 2/9/04: 11: I am not sure why
                          I need to dereference what is a full scale object
                          inside the quickBasicEngine but what is a scalar
                          .Net value when I reproduce the test in the variable
                          tester.
                          						  
02 10 2004   Nilges    1. A Date variable type is needed for better quickBasic
                          compatibility.
                          						  
02 10 2004   Nilges    1. As of today:
                          1.1 All non-UDT test cases for qbVariable work
                          1.2 No UDT test cases work
                              1.2.1 UDT,(strMember02,Array,String,1,2):("A","B")
                                    doesn't parse and produces unusable objects 
                              1.2.2 The more complex UDTs in qbVariable\Archive\
                                    02102004\udtTestCases.TXT fail in the same
                                    way
						  1.3 "Variant,(Array,Integer,1,10):32767(10)" fails
						      with the message "upperbound requested for 
						      nonarray"                                    
                          1.3 Test containment and isomorphism aren't tested 
                          1.4 Test needs to include random cases: code for end
                              of testInterface is in qbVariable\Archive\
                              02102004\randomTestCases.TXT   
						  
02 11 04   Nilges         Note that the cache code clones the code from
                          qbVariableType and ideally should be
                          common methods			
                          
02 14 04   Nilges         Status as of 3:19 AM: fails to note that comma-delimited
                          entries are beyond the end of the array   RESOLVED                   		                                
                          
02 16 04   Nilges         1. Status as of 2-16

						  1.1 Test of three dimension array fails and has been removed,
                              and archived in 3D array creation.TXT in archive\20162004.
                          
                              Need to walk through this test or a simplification to find
                              bug.                   		                                

						  1.2 UDTs are not compiled properly 
						  
						  1.3 Button on test form to produce random variables produces
						      random types and not random variables, and it produces
						      strings that cause numerous errors         
						      
04 06 04   Nilges         Documentation of VariableName property says that name must be
                          1..31 characters long. Is this true? Is this enforced?						                		                                
                                    
                          						  
                      
                      
                      
                          
                                                 