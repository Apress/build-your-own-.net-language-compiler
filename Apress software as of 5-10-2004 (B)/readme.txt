APRESS SOFTWARE FOR "BUILD YOUR OWN .NET COMPILER"


This folder contains all the software for Build Your Own .Net Compiler by Edward
G. Nilges. This readme file describes the following topics:


     *  Software included
     *  Status of this software
     *  Folder and build standards
     *  Build procedure
     *  Change and build record


Changes to the text, outside the Change and Build record, are marked with a vertical
bar.


SOFTWARE INCLUDED

This folder contains all the software mentioned in the book.

All software was written in full by Edward G. Nilges (spinoza1111@yahoo.com) and
no software has been included from any other source. It is made available under
generally understood Open Source conventions. You may use it freely in hobby and
personal applications. You may use it in corporate and shrinkwrap applications as
long as the authorship information in the source code is preserved.

The software will of course contain "issues" (a polite term for my stupid mistakes,
unpardonable omissions, and most egregious bugs). It is a work in process which as 
of today meets its goal of demonstrating how compilers work as it evolves to a more 
rugged product.

Therefore you may informally register with me for updates to the compiler by 
sending your request to spinoza1111@yahoo.COM. Place "Request for compiler updates"
in the subject line.

You are also encouraged to participate in the evolution of this software by following 
these open source conventions.


     *  When you make an improvement or a bug fix, maintain the unchanged file
        and note your changes in the changed file, using the Change Records that
        will be found, either in readme files attached to projects and solutions,
        or inline with the source code in the documentation header.

     *  Mail your fix to me at spinoza1111@yahoo.com. I will maintain an up to date
        copy of all workable fixes. 


All software with the exception of bnfAnalyzer (Chapter 4) is written in Visual
Basic .Net 2003 and may be compiled using Visual Studio Professional 2003 and
subsequent versions. To use the .Net source code with Visual Basic Standard edition,
you will have to replace all references to DLLS by references to files included
in the referencing project and all code is supplied for this.

If you are using a version of Visual Studio 2002 or prior you will probably have to
change all source code (.VB) files to use carriage return and newline in place of
newline only. This can be done by opening the file in Notepad, copying the text,
pasting the file in word, and copying the result back to a new .VB file.

To examine the source code of the Chapter 4 software bnfAnalyzer you need Visual
Studio (Visual Basic) release 6 Professional or Enterprise. You may get away using 
release 5 with modifications to the project files, but forget about using Release 4 
to examine bnfAnalyzer. If you are using the Learning edition of VB 6 you will need
to incorporate utilities as a project in the bnfAnalyzer project.

The following programs are mentioned in the book and included in this software,
all under the folder egnsf, which stands for Edward G. Nilges Software Factory.
If you have moved egnsf to a typical hard disk then all relative references below,
to egnsf, are actually C:\egnsf: you may also move the egnsf folder to a folder of
your choosing since for the most part .Net and VB-6 play nice with relative folder
references.

The following list is in alphabetic order and not dependency order; see DLL and EXE
Dependencies, below, for dependencies.


     *  bnfAnalyzer: a Visual Basic 6 application which analyzes Backus-
        Naur form files and prints a skeleton reference manual for the
        language described as seen in Chapter 4.

        To RUN bnfAnalyzer, open egnsf\bnfAnalyzer\bnfAnalyzer.EXE. Note that
        bnfAnalyzer is a COM application.

        Don't you love COM? Isn't this a .Net book?

        COM sucks, and this is a .Net book. However, when the time came to write
        the software for Chapter 4 my Vaio with .Net was in the pawnshop.

        To EXAMINE the source code of bnfAnalyzer, open egnsf\bnfAnalyzer\
        bnfAnalyzer.VBP. Use Visual Basic 6 Professional or Enterprise.
               
     *  collectionUtilities: a project to generate a set of utilities. These 
        utilities are used for common tasks using Collections.

        To EXAMINE collectionUtilities, open egnsf\collectionUtilities\
        collectionUtilities.SLN.
               
     *  creditEvaluation: a .Net application that shows how a compiler engine
        can be used to evaluate business rules as seen in Chapter 9.

        To RUN creditEvaluation, open egnsf\Apress\creditEvaluation\bin\
        creditEvaluation.EXE.

        To EXAMINE creditEvaluation, open egnsf\Apress\creditEvaluation\
        creditEvaluation.SLN.

     *  integerCalc: a very simple flyover compiler and interpreter for
        numeric calculation as seen in Chapter 3.

        To RUN integerCalc, open egnsf\Apress\integerCalc\bin\integerCalc.EXE.

        To EXAMINE the source code of integerCalc, open egnsf\integerCalc\
        integerCalc.SLN.

     *  qbGUI: a .Net application that tests the quickBasicEngine, its Nutty
        Professor interpreter, and MSIL code generation in a Graphical User 
        Interface (GUI) that shows how the compiler works, as seen in Chapters
        7 and 8.

        To RUN qbGUI, open egnsf\Apress\quickBasic\qbGUI\bin\qbGUI.EXE.

        To EXAMINE qbGUI and the full compiler, open egnsf\Apress\quickBasic\
        qbGUI\qbGUI.SLN

     *  qbScannerTest: a .Net application that tests the data model and
        class for lexical analysis, known as qbScanner, as seen in Chapter 5.

        To RUN qbScannerTest, open egnsf\Apress\quickBasic\qbScanner\qbScannerTest\
        qbScannerTest.EXE.

        To EXAMINE qbScannerTest and qbScanner, open egnsf\Apress\quickBasic\qbScanner\
        qbScannerTest.SLN.

     *  qbVariableTest: a .Net application that tests the data model and
        class for Quick Basic variables, known as qbVariable, as seen 
        in Chapter 6.

        To RUN qbVariableTest, open egnsf\Apress\quickBasic\qbVariable\
        qbVariableTest\bin\qbVariableTest.EXE.

        To EXAMINE qbVariableTest and qbVariable, open egnsf\Apress\quickBasic\
        qbVariable\qbVariableTest\qbVariableTest.SLN

     *  qbVariableTypeTester: a .Net application that tests the data model and
        class for Quick Basic variable types, known as qbVariableType, as seen 
        in Chapter 6.

        To RUN qbVariableTypeTester, open egnsf\Apress\quickBasic\qbVariableType\
        qbVariableTypeTester\bin\qbVariableTypeTester.EXE.

        To EXAMINE qbVariableTypeTester and qbVariableType, open egnsf\Apress\quickBasic\
        qbVariableType\qbVariableType.SLN

     *  relab: a laboratory for entering, testing and documenting your
        regular expressions and accessing some example regular expressions as
        seen in Chapter 3.

        To RUN relab, open egnsf\relab\bin\relab.EXE.

        To EXAMINE the source code of relab, open egnsf\relab\
        relab.SLN.

     *  utilitiesTest: a project to generate and test a basic set of Shared utilities
        including the project to generate these utilities. These utilities are used
        in other projects for utility functions independent of presentation.

        To RUN utilitiesTest, open egnsf\utilities\bin\utilitiesTest.EXE.

        To EXAMINE the source code of utilities, open utilitiesTest.VBP.

     *  windowsUtilities: a project to generate a set of Shared utilities. These 
        utilities are used in other projects for Windows presentation tasks.

        To EXAMINE the source code of windowsUtilities, open windowsUtilities.VBP.

     *  zoom: a project to generate a visual control for examining text. 

        To EXAMINE the source code of zoom, open zoom.VBP.


FOLDER AND BUILD STANDARDS

Each project contains a readme file with a mission statement, a change record, and,
as appropriate, a list of open issues and bugs. Most solutions contain a
solutionReadme file with a mission statement, change record and issue list.

The most important modules also contain quotes from literature and philosophy to set
a high cultural level.

Many projects use some or all of the following utility packages, for which object
and source have been provided.

     
     *  utilities: a static (stateless aka Shared) class providing general
        string processing and math functionality to extend Visual Basic. Source
|       code may be viewed by opening egnsf\utilities\utilities.SLN.

        The utilities may be tested by running egnsf\utilities\utilitiesTest\
        utilitiesTest.EXE. Run this executable and click its Test button on its
        lower right-hand corner. Several tests will execute, and, you should see
        a success message.

        The VB-6 utilities for the bnfAnalyzer are a class (clsUtilities) 
        embedded as utilities.CLS in the bnfAnalyzer.VBP project.

     *  collectionUtilities: a class providing a set of tools for dealing with
        Collections and extending Collections to the support of compiler data
        structures. Source code may be viewed by opening egnsf\collectionUtilities\
        collectionUtilities.SLN.

     *  windowsUtilities: a static class providing some tools for working with
        Windows applications. Source code may be viewed by opening egnsf\
        windowsUtilities\windowsUtilities.SLN.

     *  zoom: a class that synthesizes a simple, read-only form for viewing text
        information aligned using spaces rather than tabs at RTF. Source code may
        be viewed by opening egnsf\zoom\zoom.SLN.


In general but with exceptions:


     *  Projects internal to the mission of the compiler are included in solutions
        as projects and project references

     *  The above utilities are included as DLL references


The exception is collectionUtilities which is included as a project and as a project
reference in some compiler projects. 


BUILD PROCEDURE

For best results, use Visual Studio 3 Enterprise or Professional Edition as follows
to build the compiler and its Graphical User Interface from the ground up. The following
procedure will ensure that all References are current, and, it tests the compiler and
its components.

Despite the fact that Option Strict is on all modules should in this procedure compile
with no warnings or error messages.

Click Yes to all prompts that warn you that adding a reference will change where 
another reference will be found. 


     1.  Open egnsf\utilities\utilities.SLN and Build utilities.DLL

     2.  Open egnsf\collectionUtilities\collectionUtilities.SLN 

         2.1 Ensure that its reference to utilities.DLL is current (for best
             results, remove the current reference and add the reference
             again)

         2.2 Build 

     3.  Open egnsf\windowsUtilities\windowsUtilities.SLN 

         3.1 Ensure that its reference to utilities.DLL is current

         3.2 Build 

     4.  Open egnsf\zoom\zoom.SLN 

         4.1 Ensure that the references to the following are current.
             4.1.2 utilities.DLL
             4.1.3 windowsUtilities.DLL

         4.2 Build 

     5.  Open egnsf\utilities\utilitiesTest\utilitiesTest.SLN 

         5.1 Ensure that the references to the following are current.
             5.1.1 collectionUtilities.DLL
             5.1.2 utilities.DLL
             5.1.3 windowsUtilities.DLL
             5.1.3 zoom.DLL

         5.2 Build 

         5.3 Run egnsf\utilities\utilitiesTest\Bin\utilitiesTest.EXE.

             You shall see (assuming you haven't run it before) an
             About screen. Dismiss this screen to see procedures load,
             unnecessarily (the option to do so supports documentation
             and can be suppressed). 

             Click the Test button and expect to see a success message.
             If you want, you can review the results in detail.

     6.  Open egnsf\Apress\quickBasic\qbTokenType\qbTokenType.SLN 

         6.1 Build 

     7.  Open egnsf\Apress\quickBasic\qbToken\qbToken.SLN 

         7.1 Ensure that the references to the following are current.
             7.1.1 qbTokenType.DLL
             7.1.2 utilities.DLL

         7.2 Build 

     8.  Open egnsf\Apress\quickBasic\qbScanner\qbScanner.SLN 

         8.1 Ensure that the references to the following are current.
             8.1.1 collectionUtilities.DLL
             8.1.2 qbToken.DLL
             8.1.3 qbTokenType.DLL
             8.1.4 utilities.DLL

         8.2 Build 

     9.  Open egnsf\Apress\quickBasic\qbScanner\qbScannerTest\qbScannerTest.SLN 

         9.1 Ensure that the references to the following are current.
             9.1.1 collectionUtilities.DLL
             9.1.2 qbScanner.DLL
             9.1.3 qbToken.DLL
             9.1.4 qbTokenType.DLL
             9.1.5 utilities.DLL
             9.1.6 windowsUtilities.DLL
             9.1.7 zoom.DLL

         9.2 Build 

         9.3 Run egnsf\APress\quickBasic\scanner\qbScannerTest\Bin\qbScannerTest.EXE.

             You shall see (assuming you haven't run it before) an
             About screen. Dismiss this screen. 

             Click the Test button and expect to see a success message.
             If you want, you can review the results in detail.

     10. Open egnsf\Apress\quickBasic\qbVariableType\qbVariableType.SLN 

         10.1 Ensure that the references to the following are current.
              10.1.1 collectionUtilities.DLL
              10.1.2 qbScanner.DLL
              10.1.3 qbTokenType.DLL
              10.1.4 utilities.DLL

         10.2 Build 

     11. Open egnsf\Apress\quickBasic\qbVariableType\qbVariableTypeTester\qbVariableTypeTester.SLN 

         11.1 Ensure that the references to the following are current.
              11.1.1 qbVariableType.DLL
              11.1.2 utilities.DLL
              11.1.3 windowsUtilities.DLL
              11.1.4 zoom.DLL

         11.2 Build 

         11.3 Run egnsf\APress\quickBasic\qbVariableType\qbVariableTypeTester\Bin\qbVariableTypeTester.EXE.

              You shall see (assuming you haven't run it before) an
              About screen. Dismiss this screen. 

              Click the Stress Test button and expect to see (1) a series of 
              progress reports and then (2) a success message. If you want, 
              you can review the results in detail.

     12. Open egnsf\Apress\quickBasic\qbVariable\qbVariable.SLN 

         12.1 Ensure that the references to the following are current.
              12.1.1 collectionUtilities.DLL
              12.1.2 qbScanner.DLL
              12.1.3 qbTokenType.DLL
              12.1.4 qbVariableType.DLL
              12.1.5 utilities.DLL

         12.2 Build

     13. Open egnsf\Apress\quickBasic\qbVariable\qbVariableTest\qbVariableTest.SLN 

         13.1 Ensure that the references to the following are current.
              13.1.1 qbVariable.DLL
              13.1.2 qbVariableType.DLL
              13.1.3 utilities.DLL
              13.1.4 windowsUtilities.DLL
              13.1.5 zoom.DLL

         13.2 Build

         13.3 Run egnsf\APress\quickBasic\qbVariable\qbVariableTest\Bin\qbVariableTest.EXE.

              You shall see (assuming you haven't run it before) an
              About screen. Dismiss this screen. 

              Click the Stress Test button and expect to see (1) a series of 
              progress reports and then (2) a success message. If you want, 
              you can review the results in detail.

     14. Open egnsf\Apress\quickBasic\qbOp\qbOp.SLN 

         14.1 Ensure that the references to the following are current.
              14.1.1 utilities.DLL

         14.2 Build

     15. Open egnsf\Apress\quickBasic\qbPolish\qbPolish.SLN 

         15.1 Ensure that the references to the following are current.
              15.1.1 qbOp.DLL
              15.1.2 qbVariable.DLL
              15.1.3 utilities.DLL

         15.2 Build

     17. Open egnsf\Apress\quickBasic\quickBasicEngine\quickBasicEngine.SLN 

         17.1 Ensure that the references to the following are current.
              17.1.1 collectionUtilities.DLL
              17.1.2 qbOp.DLL
              17.1.3 qbPolish.DLL
              17.1.4 qbScanner.DLL
              17.1.5 qbToken.DLL
              17.1.6 qbTokenType.DLL
              17.1.7 qbVariable.DLL
              17.1.8 qbVariableType.DLL
              17.1.9 utilities.DLL

         17.2 Build


     18. Open egnsf\Apress\quickBasic\quickBasicEngine\quickBasicEngine.SLN 

         18.1 Ensure that the references to the following are current.
              18.1.1 collectionUtilities.DLL
              18.1.2 qbOp.DLL
              18.1.3 qbPolish.DLL
              18.1.4 qbScanner.DLL
              18.1.5 qbToken.DLL
              18.1.6 qbTokenType.DLL
              18.1.7 qbVariable.DLL
              18.1.8 qbVariableType.DLL
              18.1.9 utilities.DLL

         18.2 Build

         18.3 Run egnsf\APress\quickBasic\qbGUI\Bin\qbGUI.EXE.

              You shall see (assuming you haven't run it before) an
              About screen. Dismiss this screen. 

              Click the More button (to expand the screen) and then
              click the Test button and expect to see (1) a series of 
              progress reports and then (2) a successful test report. If you want, 
              you can review the results in detail.





C H A N G E   A N D   B U I L D   R E C O R D --------------------------------------
  DATE     PROGRAMMER     DESCRIPTION AND EVENTS
--------   ----------     ----------------------------------------------------------
01 24 04   Nilges         Shipped a build to Dan Appleman after testing all projects
             
                          Dan found a build problem in qbScannerTest: missing refs
                          to collectionUtilities and utilities.

01 25 04   Nilges         Added this log

01 26 04   Nilges         1.  Changes from Chapter 4 review (see also individual
                              logs)

                              1.1 The VB-6 utilities class has been removed and placed
                                  inside the only project which needs it, bnfAnalyzer.

                              1.2 bnfAnalyzer has been changed for richer output in
                                  support of Chapter 4's text: see Change Record in
                                  bnfAnalyzer.FRM.

02 02 04   Nilges         Shipped a build to Apress after testing all projects

02 02 04   Nilges         1.  Changes from Chapter 5 review (see also individual
                              logs)

                              1.1 Various bnfAnalyzer improvements and bug fixes.

                              1.2 Various relab improvements and fixes.

                              1.3 Various qbScanner and qbScannerTest improvements 
                                  and fixes.


02 16 04   Nilges         1.  Changes from Chapter 6 review (see also individual
                              logs)

03 21 04   Nilges         Shipped a build to Apress after testing all projects

05 05 04   Nilges         1.  Updated content
                          2.  Added Build Procedure and Dependencies sections

