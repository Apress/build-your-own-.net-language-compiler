qbScanner

This class scans input source code for the quickBasic engine and provides, 
on demand, scanned source tokens and scanned lines of source code. This class
uses "lazy" evaluation, scanning the source code only when necessary, and when 
an unparsed token is requested.

C H A N G E   R E C O R D -------------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION
--------   ----------     -------------------------------------------------
04 30 03   Nilges         Started version 1.0
                          
05 07 03   Nilges         Completed version 1.0 and archived to
                          C:\EGNSF\APRESS\QUICKBASIC\QBSCANNER\
                          ARCHIVE\05072003\VERSION1 
                          
05 12 03   Nilges         Bug: attempt to convert Nothing in 
                          tokenArray2String         
                          
06 21 03   Nilges         Version 2.0

                          1.  Added the test method and the TestString
                              property         
                              
						  2.  Bug: inspect did not return false or test
						      results outside bounds
						      
						  3.  Need to use XML and not vbExpression for
						      toString			
						      
08 18 03   Nilges         Version 2.0 has been archived to C:\EGN\EGNSF\
                          Apress\quickBasic\qbScanner\Archive\081803\
                          Version2					
                          
06 21 03   Nilges         Version 2.1

                          1.  Add inspection for full scan		
                          
                          2.  Made ICloneable
                          
                          3.  Made ICompareable
                          
                          4.  Added normalize method
                          
                          5.  Added token method
                          
                          6.  Added token index and count to the
                              scanEvent
                              
                          7.  Added the findRightParenthesis method 
                              
                          8.  Added the checkToken method 
                          
                          9.  Added the isInteger method
                          
10 01 03   Nilges         1.  Uploaded, to Apress FTP site                          
                          
10 02 03   Nilges         1.  checkToken is case-insenstive    

10 08 03   Nilges         1.  Bug: failed to scan an isolated period
                              properly: added the period as a token                      

10 08 03   Nilges         1.  The intIndex parameter of findRightParenthesis
                              is now ByVal because it isn't modified
                              
						  2.  Added sourceMid method
						  
						  3.  Added methods to access token properties
						  
						      2.1 tokenStartIndex
						      2.2 tokenEndIndex
						      2.3 tokenLength
						      2.4 tokenLineNumber 	    
						      2.5 tokenType 
						      
						  4.  Bug: the usability check did not return
						      False when the object wasn't usable    
						      
						  5.  Added checkTokenByTypeName method       
						  
10 15 03   Nilges         1.  Added findToken and findTokenByTypeName
                              methods
                              
10 19 03   Nilges         1.  Scanner bugs found

                              1.1 Don't set the next token index if the
                                  next token has zero length: isolated
                                  periods parse as real numbers with zero
                                  length     
                                  
11 25 03   Nilges         1.  Bug: scanned indicator not set when last
                              token past end of string                                                           
                                  
11 30 03   Nilges         1.  Bug: did not scan floating-point numbers that
							  commence with a decimal point 
                                  
12 20 03   Nilges         1.  Added dispose overlay to avoid inspection. 
                                  
02 03 04   Nilges         1.  Bug: sourceMid should return a null string for a
                              string beyond the end of the source
                              
                          2.  Added parameter check to sourceMid     
                                  
02 03 04   Nilges         1.  In scanning strings, treat left and right smart
                              quotes the same as straight quotes
                                  
02 22 04   Nilges         1.  Removed the findRightParenthesis broken feature,
                              to initialize level to one or zero depending on
                              whether a parenthesis starts the string: this
                              breaks the compiler when two left parentheses
                              are adjacent.                            
                                  
03 24 04   Nilges         1.  Removed documentation of defunct checkNumber
                              procedure.          
                              
03 31 04   Nilges         1.  Removed the documentation of findRightParenthesis
                              broken feature (cf. change record 2-22-04)        
                              
04 04 04   Nilges         1.  Removed documentation of nonexistent property
                              startIndexToken                                                                      
							  
I S S U E S -------------------------------------------------------------------
  DATE     POSTER    FIXED BY  FIX DATE D E S C R I P T I O N            
-------- ---------- ---------- -------- ---------------------------------------   
11 22 03 Nilges     Nilges     01 12 04 Was not able to test 11/30 change
										by rebuilding qbScannerTest: obtained
										build conflicts: probably need to change
										to project references exclusively
							                                                            
02 25 04 Nilges     Nilges              ampersandSuffix as a token type is not
                                        needed: it may be (should be) removed.
                              
03 22 04   Nilges                       1.  Allow tabs in addition to blanks 
                                            as white space per reference manual 
                                            
                                            Currently tabs are not allowed
                              
04 04 04   Nilges                       1.  Using the token method forces full
                                            scan
							                                                            
