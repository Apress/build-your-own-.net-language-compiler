utilitiesTest

This application and form allows the user to learn and test utilities functions.

The source code for the utilities.VB .Net library of stateless utilities for
Windows and Web applications, the windowsUtilities.VB .Net library for Windows,
the utilities.CLS library for COM, or a user library is loaded and partially
parsed to obtain its procedures, which are then displayed along with usage 
information. The usage information shows, for each procedure, which procedures
it uses (directly or indirectly), and which procedures use it, directly or
indirectly.

Another feature allows the Addison Wesley book text, in the form of a .TXT file,
to be compared against the parsed utilities files.

Also, the "lookaside" feature can store the parsed information and the usage
information on a text file and reload it when it postdates the source code. This
saves time in repeatedly parsing code and making usage analyses.
        
        
C H A N G E   R E C O R D -----------------------------------------------------
  DATE      PROGRAMMER    DESCRIPTION
--------    ----------    -----------------------------------------------------
05 10 03    Nilges        Started version 1

05 26 03    Nilges        Completed version 1: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\05262003\Version1

05 26 03    Nilges        Started version 2

                          1.  DISPLAYONLY option
                          2.  Added displayer form
                          3.  Navigation
                          4.  VB-6 support
                          5.  Usage analysis

05 26 03    Nilges        Completed version 2: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\05272003\Version2

05 26 03    Nilges        Completed version 2.1: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\05272003\Version21
                          
                          1.  Bug: off by one scanning caused adjacent procedures
                              to be missed
                          2.  Improved iteration progress report
                          3.  Reduced size of usage analysis output    

05 29 03    Nilges        1.  Renamed utilitiesInfo 
                          2.  Do not box output in the abstract window

05 26 03    Nilges        Completed Version 2.2: archived to C:\egnsf\utilities\
                          utilitiesInfo\Archive\05292003\Version22 
                          
                          1.  Scan for compound identifiers using dot notation:
                              this avoids bogus references    
                          2.  Added ability to control status level    
                          3.  Removed unused testControl

05 31 03    Nilges        1.  Renamed utilitiesTest
                          2.  Removed the utilitiesInfo form: using zoom
                          3.  Added interface to test method
                          4.  Privates and Friends need to be included but not
                              displayed for a complete usage analysis

06 02 03    Nilges        Completed Version 3: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\06032003\Version3 

06 02 03    Nilges        Completed Version 3.1: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\06022003\Version31
                          
                          1.  Optionally, don't include Private and Friend 
                              procedures in the usage lists
                              
                          2.  Usage lists need to be separated by commas and spaces
                              and not just commas    

06 02 03    Nilges        Completed Version 3.2: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\06022003\Version32
                          
                          1.  Bug: missing yes and no in message box

06 08 03    Nilges        Completed Version 3.3: archived to C:\egnsf\utilities\
                          utilitiesTest\Archive\06082003\Version33
                          
                          1.  Added zoom reference manual button

06 08 03    Nilges        Version 3.4
                          
                          1.  Ability to match text with results
                          2.  Updated About information
                          3.  Sort the procedures
                          4.  Lookaside
                          
12 06 03    Nilges        Added item testing form                          
                          
12 12 03    Nilges        Bug: need to create collection utilities                         
                          
O P E N   I S S U E S -------------------------------------------------------------
  DATE     ISSUE AUTHOR   TYPE OF ISSUE   ISSUE DESCRIPTION
--------   ------------   -------------   -----------------------------------------
06 11 03   Nilges         Minor           The code contains parsing and storage of
                                          procedure parameters, yet this information
                                          is not used.
                                          
                                          This anticipated the need to use this
                                          information and it has been left for 
                                          a future implementation of an automated
                                          testing GUI which can construct test
                                          screens on the fly. 
                                          
                                          In particular, note that the lookaside
                                          feature doesn't save and restore this
                                          parameter information.                           

04 09 04   Nilges         Minor           Form not adjusted to available screen.
                               

