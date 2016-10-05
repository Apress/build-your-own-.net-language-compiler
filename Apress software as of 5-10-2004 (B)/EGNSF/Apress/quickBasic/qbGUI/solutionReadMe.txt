This Solution creates the qbGUI application which allows you to create, edit and
run Quick Basic programs while watching how the scanner, compiler, assembler and
Nutty Professor interpreter all work together.


C H A N G E   R E C O R D (changes and events after 1-8-2004) ---------
  DATE     PROGRAMMER     DESCRIPTION OF CHANGE
--------   ----------     ---------------------------------------------	  
01 09 04   Nilges         Delivered current version to Apress on this
                          date
                          
01 12 04   Nilges         Replaced References to utilities to DLLs
                          
02 20 04   Nilges         1.  Form changes for tooltips, error message
                              display
                          2.  Bug occured in variableType object and was
                              fixed    
                          3.  Bug occured in qbOp object and was
                              fixed    
                          
03 01 04   Nilges         1.  Changed to use DLL rather than project
                              includes upon release to Apress    


I S S U E S -----------------------------------------------------------
  DATE       POSTER       DESCRIPTION AND RESOLUTION
--------   ----------     ---------------------------------------------
02 09 04   Nilges         A minor but consistent bug exists in all
                          objects except qbVariable.

                          In their New methods, the sequence number
                          is incremented properly using an Interlocked
			              method. However, it's then used outside of
                          locking to construct the object name.

		                  This COULD produce duplicate names. Because
                          the object name is never used to uniquely
                          key the object, this bug is minor
                          
03 02 04   Nilges         1.  Add AndAlso and OrElse opcodes

                          2.  String limitations should be enforced by
                              a qbVariable property and conditioned on
                              the EXTENSION compiler flag        
                              
03 10 04                  Goals as of 3-10 for qbGUI solution

                          1.  Fix hang in parse
                          2.  Must compile helloWorld and nFactorial and
                              pass test
                          3.  Add internal resize logic    
                                                                                              
                                                