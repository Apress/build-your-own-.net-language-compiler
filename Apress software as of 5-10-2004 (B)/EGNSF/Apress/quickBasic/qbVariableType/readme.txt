qbVariableType

This class represents the type of a quickBasicEngine variable, including 
support for an unknown type.


C H A N G E   R E C O R D ---------------------------------------------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     ---------------------------------------------
04 29 03   Nilges         1.  Version 1 archived in abymeOiseaux at  
                              c:\egnsf\aPress\qbVariableType\archive 
                              \04292003\version 1     
                                            
						  2.  Added BoundSize

06 23 03   Nilges		  Version 2 archived to abymeDesOiseaux\\
                          egnsf\aPress\qbVariableType\archive\
                          06252003\version 2
						  1.  Added className
						  2.  Added defaultValue 
                          3.  Added netValueInQBdomain	
						  4.  Added netValue2QBdomain	
						  5.  Added netType2QBdomain	
						  6.  Added qbDomain2NetType
						  7.  Bug: need to check for null fromString
						  8.  Added QBVARIABLETYPE_TEST and
						      TestAvailable
						  9.  Added isArray 
						  
06 25 03   Nilges         1.  Allow derivation of defaultValue for
                              instantiated type						     
						  
06 26 03   Nilges         1.  Added isUnknown						     
						  
07 16 03   Nilges         1.  Make About display narrow in object2XML
						  
07 25 03   Nilges         1.  Bug: narrowerType does not handle Booleans
                              properly wrt to Bytes
                              
                          2.  Renamed narrowerType to be containedType  
                              
                          3.  Syntax needs to support polymorphous variant
                              arrays	
                              
						  4.  Added mkType method	                     
						  
						  5.  Added toName method
						  
						  6.  Added mkRandomFromString and mkRandomArray        
						  
						  7.  Allow Variants to contain Unknown for future
						      use
						  
						  7.  Added the testEvent
						  
						  8.  Return error and/or usability information in
						      toString
						      
						  9.  Place toString in inspect
						  
						  10. Added explanation to compareTo method
						  
						  11. Added the Tag property
						  
08 24 03    Nilges        1.  Made isScalarType public and overloaded it
                              with a string argument
                              
                          2.  Added netValue2QBvalue
                          
                          3.  Added mkRandomDomain    						  
                          
                          4.  Added mkRandomScalar
                          
                          5.  Bug: fromString assigned an invalid default
                              for arrays
                              
                          6.  Added mkRandomScalarValue    
                          
                          7.  Added mkRandomVariantValue
                          
                          8.  Added arraySlice
                          
                          9.  Restrict mkRandomArray dimensions to 1..3,
                              and bound size to 20, since otherwise array
                              is too large.
                              
                          10. Added isDefaultValue method
                          
08 30 03    Nilges        1.  Bug: isScalar needs to test not unknown, variant,
                              or Null                          
                          
                          2.  Added toDescription method
                          
                          3.  Bug: mkRandomType returned scalars only
                          
09 30 03    Nilges        1.  Added support for UDTs

                              1.1 Modifies the following procedures
                              
                                  1.1.1 containedType      
                                  1.1.2 defaultValue
                                  1.1.3 fromString (toString/fromString 
                                        syntax expands to include variable 
                                        name and semicolon-separated list 
                                        of types: fromString now uses the
                                        qbScanner to scan)
                                  1.1.4 hungarianPrefix      
                                  1.1.5 innerType syntax expands, is now a
                                        method
                                  1.1.6 inspect
                                  1.1.7 isUDT added
                                  1.1.8 UDTmemberName and UDTmemberCount
                                  1.1.9 mkRandomType needs to select UDT
                                        in random cases
                                  1.1.10 mkRandomUDT added  
                                  1.1.11 object2XML  
                                  1.1.12 StorageSpace  
                                  1.1.13 string2enuVarType  
                                  1.1.14 test
                                  1.1.15 toContainedName
                                  1.1.16 toDescription
                                  1.1.17 toString     
                                  1.1.18 varType is now a method
                                  
                              1.2 New state: objVariableType now may be a
                                  variableType or a Collection of variable
                                  types  

							  1.4 8/30 version archived, on abymedesoiseaux
							      machine, to C:\EGNSF\APRESS\QUICK BASIC\
							      QBVARIABLETYPE\ARCHIVE\09302003\
							      Version as of 08302003  
							      
10 04 03   Nilges          Bug: did not parse nested UDTs

10 06 03   Nilges          Removed street address from About

10 11 03   Nilges          Removed need for qbToken reference	

                           Added error location to parse error message						                                

12 20 03   Nilges          Added booInspect parameter to dispose method	

01 12 04   Nilges          Don't use fromString parse for the clone

                           1. Test time before this change in a compiled
                              version: 22 seconds	

02 07 04   Nilges          1. Made the object IDisposable DONE
                           2. Added disposeInspect DONE
                           3. Control width of object2XML in test DONE
                           4. Add a cache DONE
                              4.1 Test method timing when not used: 16 seconds
                                  in a standard test, compiled version
                              4.2 Timing when used, size 100: 5 seconds
                                  in a standard test, compiled version
							  4.3 Compiler impact: evaluation of 25! using
							      nFactorial.BAS in the qbGUI with "Full Monty" 
							      display takes 58 secs without cacheing: 
							      it takes only 40 secs with cacheing.                                  
							  4.3 Business rules impact: evaluation of standard
							      scenario takes 33 secs without cacheing: 
							      it takes only 25 secs with cacheing.                                  
                           5. Made this object IDisposable DONE  
                           6. Output cache state to XML DONE
                           7. Be able to cache UDTs with inner vts DONE

02 10 04   Nilges          1. netValue2QBdomain changed to return the QB domain
                              Boolean when the .Net value is Boolean

02 11 04   Nilges          1. Added object2Type method
						   2. Added fromString2TypeName method
						   3. Added changeArrayBound method
						   4. Added validFromstring method
						   5. Made LowerBound and UpperBound properties
						      read-write
						   6. Added redimension method   

02 14 04   Nilges          1. Bug: UDTmember properties failed to return
                              Nothing on an error
                              1.1 Added Return Nothing 

02 15 04   Nilges          1. Bug: the container type for a Variant can
                              be Something (Null) when the variant is
                              abstract
                              1.1 Added code
                              
02 22 04   Nilges          1. Bug: when I ran this integrated with the
                              Quick Basic GUI, I received an error message:
                              Variant, Null was considered abstract
                              1.1 Returning True ONLY when the variant's
                                  contained type is Nothing and not when
                                  this type is Null. Retested and the
                                  code passes qbVariableType and qbVariable
                                  stress tests.
                              
02 29 04   Nilges          1. Added the BoundList property
                              
04 06 04   Nilges          1. Documentation corrections
                              
05 01 04   Nilges          1. Bug: name2NetType gave incorrect result for
                              "object"
                              1.1 Fixed
                          
I S S U E S ------------------------------------------------------------
  DATE        POSTER      I S S U E   D E S C R I P T I O N   RESOLUTION
--------   ------------   ---------------------------------   ----------
08 30 03   Nilges         Poor integration with qbVariable:   9/30 change                         
                          rescans the type fromString,        addresses  
                          using own scanner
                          
09 30 03   Nilges         Can't test UDTs for either
                          containment or isomorphism       
                          
10 03 03   Nilges         toDescription produces confusing
                          "total size" messages for
                          nested types        
                          
10 11 03   Nilges         Parse error message still sucks                                              

02 08 04   Nilges         1. Nondeterministic tests MAY
                             fail when testing type con-
                             tainment on an invalid index
                          2. Use State assignment to clone
                          3. Expose CacheSize to control
                             size of cache
                             
02 11 04   Nilges         object2Type should probably return
                          Array for an orthogonal collection
                          and UDT for a non-orthogonal 
                          collection...if this is useful.
                          
                          Currently it will return Unknown. 
                                                      
02 11 04   Nilges         Need identified wrt both to
						  qbVariableType and qbVariable for
						  a qbVariableCalculus Shared class 
						  to overload the common arithmetic
						  (and, perhaps, the logical and
						  compare) operators for fromString
						  expressions.
						  
						  This is because the qbVariable
						  parser needs to attach (sum?)
						  qbVariables and their types.
						  
						  This need will be handled at present
						  by code in the qbVariable parser.
                                                      
02 11 04   Nilges         cf. Note in documentation of clone
                          method: test fast clone and,
                          ultimately, retire the slow clone.
                          
04 16 04   Nilges         mkRandomScalar has an overload which
                          allows it to be misused, since it may
                          be passed a strTypes parameter which
                          specifies non-scalar types                          
