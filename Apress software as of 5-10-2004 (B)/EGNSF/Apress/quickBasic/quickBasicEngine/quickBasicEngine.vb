Option Strict

Imports System.Threading
Imports qbOp.qbOp
Imports System.Reflection
Imports System.Reflection.Emit
Imports DotNetAssembly = System.Reflection.Assembly

' *********************************************************************
' *                                                                   *
' * quickBasicEngine     quickBasicEngine                             *
' *                                                                   *
' *                                                                   *
' * This class does all scanning, parsing and interpretation for my   *
' * version of Quick Basic. It may be "dropped in" to a .Net          *
' * application and it will provide the ability to evaluate immediate *
' * Basic expressions as well as compile and run Basic programs.      *
' *                                                                   *
' * The rest of this comment block describes the following topics.    *
' *                                                                   *
' *                                                                   *
' *      *  The quickBasicEngine data model                           *
' *                                                                   *
' *      *  The quickBasicEngine state                                *
' *                                                                   *
' *      *  Object availability states: usable/unusable: ready/       *
' *         running(N)/stopping/stopped                               *
' *                                                                   *
' *      *  Properties, methods, and events of this class             *
' *                                                                   *
' *      *  Multithreading considerations                             *
' *                                                                   *
' *      *  Compile-time symbols                                      *
' *                                                                   *
' *      *  Error taxonomy                                            *
' *                                                                   *
' *      *  References                                                *
' *                                                                   *
' *                                                                   *
' * THE QUICKBASICENGINE DATA MODEL --------------------------------- *
' *                                                                   *
' * The state of this class consists of all source code for a program,*
' * its scanned representation, and the names and the structured      *
' * values of all variables found in the code as well as details,     *
' * described in the next section.                                    *
' *                                                                   *
' * Note that no explicit "parse tree" is built because this is       *
' * unnecessary; parse information is available just-in-time through  *
' * the parseEvent.                                                   *
' *                                                                   *
' *                                                                   *
' * THE QUICKBASICENGINE STATE -------------------------------------- *
' *                                                                   *
' * The state of the quickBasicEngine consists of the following.      *
' *                                                                   *
' *                                                                   *
' *      *  booUsable: the object usability switch                    *
' *                                                                   *
' *      *  objThreadStatus: contains the thread status of the object *
' *         as a Private miniobject of type threadStatus. See the     *
' *         discussions below on "object availability states" and     *
' *         "multithreading considerations" for more information about*
' *         thread status.                                            *
' *                                                                   *
' *      *  strName: the object instance name                         *
' *                                                                   *
' *      *  objScanner: the just-in-time scanner including the source *
' *         code                                                      *
' *                                                                   *
' *      *  colPolish: the collection, of qbPolish objects, generated *
' *         for the source code by the quickBasicEngine's parser.     *
' *         This is an unkeyed collection.                            *
' *                                                                   *
' *      *  colVariables: the collection, of qbVariable objects,      *
' *         generated for the source code by the quickBasicEngine's   *
' *         parser.                                                   *
' *                                                                   *
' *         This collection is keyed using the name of the variable.  *
' *         The Tag property of each member is set to the index of the*
' *         variable object in the collection.                        * 
' *                                                                   *
' *      *  objCollectionUtiities: A collectionUtilities object to    *
' *         support collections                                       *
' *                                                                   *
' *      *  booAssembled: indicates whether the code has been         *
' *         assembled in addition to being compiled                   * 
' *                                                                   *
' *      *  booCompiled: indicates whether the code has been          *
' *         compiled                                                  * 
' *                                                                   *
' *      *  enuSourceCodeType: indicates the type of the source code: *
' *                                                                   *
' *         + Unknown: not set                                        *
' *         + Immediate: immediate expression for immediate evaluation*
' *         + Program: full-scale source program                      *
' *                                                                   *
' *      *  booConstantFolding: indicates whether compile time        *
' *         "folding" (evaluation) is in effect                       * 
' *                                                                   *
' *      *  booDegenerateOpRemoval: indicates whether compile time    *
' *         removal of "degenerate" operations including addition of  *
' *         zero and multiplication by one is in effect               * 
' *                                                                   *
' *      *  objImmediateResult: saves the result of the most recent   *
' *         operation (as a qbVariable)                               *
' *                                                                   *
' *      *  booExplicit: indicates whether the current source code    *
' *         is a full program containing the Option Explicit          *
' *         directive                                                 *
' *                                                                   *
' *      *  usrSubFunction: table of subroutines and functions: each  *
' *         entry contains:                                           *
' *                                                                   *
' *         + booFunction: True for function, False for subroutine    *
' *                                                                   *
' *         + strName: name of the procedure                          *
' *                                                                   *
' *         + usrFormalParameters: sublist of the formal parameters   *
' *           declared in the procedure header: each entry contains:  * 
' *                                                                   *
' *           - strName: formal parameter name                        *  
' *                                                                   *
' *           - booByVal: True (parameter is passed by value) or False*
' *             (parameter is passed by reference)                    *
' *                                                                   *
' *         + intLocation: location of the start of the procedure in  *
' *           the colPolish collection or 0 for an undefined procedure*
' *                                                                   *
' *      *  colSubFunctionIndex: index to table of subroutines and    *
' *         functions                                                 *
' *                                                                   *
' *      *  intLabelSeq: sequence number for assembler labels         *
' *                                                                   *
' *      *  usrConstantExpression: table of recent constant           *
' *         expressions: used when the ConstantFolding property needs *
' *         to "lazily" avoid evaluation of a constant or constant    *
' *         expression.                                               *
' *                                                                   *
' *         The oldest constants and constant expressions in this     *
' *         table are purged when the table reaches a maximum size.   *
' *                                                                   *
' *         Each entry contains the following fields:                 *
' *                                                                   *
' *         + strConstantExpression: the constant or expression       *
' *         + objValue: the value of the constant as a qbVariable     *
' *                                                                   *
' *      *  colConstantExpressionIndex: the index of constant         *
' *         expressions. Each entry is a subcollection with two items.*
' *         item(1) is the constant expression and item(2) is its     *
' *         index in usrConstantExpression.                           *
' *                                                                   *
' *      *  intOldestConstantIndex: pointer to the oldest entry in the*
' *         constant table                                            *
' *                                                                   *
' *      *  colReadData: queue of Read Data inputs                    *
' *                                                                   *
' *      *  intReadDataIndex: next read data index                    *
' *                                                                   *
' *      *  colLabel: label collection. Key is symbolic label or      *
' *         numeric label preceded by underscore: data is a two-item  *
' *         subcollection that contains the key in Item(1) and the    *
' *         label's Polish location in Item(2).                       *
' *                                                                   *
' *      *  booAssemblyRemovesCode: True if assembly should remove    *
' *         comments and labels, False otherwise.                     *       
' *                                                                   *
' *      *  colEventLog: log of events raised                         *       
' *                                                                   *
' *      *  booGenerateNOPs: this state variable is generated only    *
' *         when the QUICKBASICENGINE_EXTENSION compile-time flag is  *
' *         present and True: it will itself be True when the compiler*
' *         should generate NOP and REM commands in the object code,  *
' *         False otherwise                                           *
' *                                                                   *
' *                                                                   *
' * OBJECT STATES---------------------------------------------------- *
' *                                                                   *
' * At any time, instances of this class are usable or nonusable, and *
' * one of Ready to run, Running (n threads), Stopping or Stopped.    *
' *                                                                   *
' *                                                                   *
' *      *  When the object is fully initialized, and after normal    *
' *         procedures have terminated, it is Ready to run.           *
' *                                                                   *
' *      *  When the object is running normally it is Running. It     *
' *         may be simultaneously running procedures in more than one *
' *         thread.                                                   *
' *                                                                   *
' *      *  When the user has requested a stop, through the stopQBE   *
' *         method, but threads are still running, the object is      *
' *         Stopping.                                                 *
' *                                                                   *
' *      *  When the user has requested a stop, through the stopQBE   *
' *         method, and no threads are still running, the object is   *
' *         Stopped.                                                  *
' *                                                                   *
' *                                                                   *
' * The usability and running states, and the number of running       *
' * threads are available through methods and properties.             *
' *                                                                   *
' * At the end of a successful New constructor the instance becomes   *
' * usable and Ready to run, whereas a successful or failed execution *
' * of dispose makes the object unusable.                             *
' *                                                                   *
' * A serious internal error: failure to create a resource: or "object*
' * abuse" (using the object after a serious error, externally        *
' * reported) makes the object not usable. The mkUnusable method may  *
' * also be used to force the object into the unusable state, and the *
' * Usable property tells the caller whether the object is usable.    *
' *                                                                   *
' * Usability makes run status moot because an Unusable instance won't*
' * run. It should be disposed, a new instance should be created, and *
' * the processing that created the error shouldn't be repeated.      *
' *                                                                   *
' * The stopQBE method places the object in a stopping state immediat-*
' * ely, and it puts the object in a stopped state when all running   *
' * threads have terminated. While an Unusable object cannot be made  *
' * Usable, a Stopped object can be restored to active duty using the *
' * resumeQBE method.                                                 *
' *                                                                   *
' * When the object instance is Stopped the state of the engine       *
' * becomes immediately Stopping as an atomic operation. Then:        *
' *                                                                   *
' *                                                                   *
' *      *  Any executing For or Do loop in the engine, which issues  *
' *         loop events, is exited as soon as the loop event is       *
' *         issued.                                                   *
' *                                                                   *
' *      *  If the interpreter is running the interpreter is exited   *
' *         after an undefined amount of time, after the object is    *
' *         stopped. The time is the duration from the issuance of    *
' *         the stopQBE and the point at which the interpreter        *
' *         arrives at the head of its for loop.                      *
' *                                                                   *
' *      *  No Public procedure will execute, and all Public          *
' *         procedures will return default values, until Stopped is   *
' *         set to False.                                             *
' *                                                                   *
' *                                                                   *
' * The resumeQBE method places the object in the Ready state if the  *
' * object is stopped. The resume method has no effect when the object*
' * is already Ready, or is in the Running or Stopping states.        *
' *                                                                   *
' * The getThreadStatus method returns the run status as one of ready,*
' * running, stopping or stopped.                                     *
' *                                                                   *
' * When the Quick Basic engine is Running, the runningThreads        *
' * method returns the number of threads that are running methods and *
' * properties inside the Quick Basic engine as a number between 1 and*
' * n. When the Quick Basic engine is Stopped or Ready,               *
' * runningThreads returns 0.                                         *
' *                                                                   *
' *                                                                   *
' * PROPERTIES, METHODS, AND EVENTS OF THIS CLASS ------------------- *
' *                                                                   *
' * Properties of this class start with an upper case letter; methods *
' * and events with a lower case letter; events end in Event.         *
' *                                                                   *
' * Any method, which does not otherwise return a value, will return  *
' * True on success or False on failure.                              *
' *                                                                   *
' * Note that this object is ICloneable and IComparable; see the      *
' * clone method and the compareTo method                             *
' *                                                                   *
' *                                                                   *
' * About: this shared, read-only property returns information about  *
' *      this class                                                   *
' *                                                                   *
' * assemble: this method assembles the Polish tokens, replacing      *
' *      symbolic labels with numeric addresses and removing comment  *
' *      lines inserted by the compiler. Assembly is a two-pass       *
' *      process: pass one converts the dictionary mapping labels     *
' *      to addresses, and pass two replaces the labels.              *
' *                                                                   *
' * assembled: this method returns True if the current source code has*
' *      been assembled, False otherwise                              *
' *                                                                   *
' * AssemblyRemovesCode: this read-write property returns and may be  *
' *      set to True to remove REMarks as inserted by the compiler,   *
' *      and label statements, during assembly.                       *
' *                                                                   *
' *      By default, this removal occurs. Note that setting this      *
' *      property to False does not change the effect of Quick Basic  *
' *      code, only its efficiency.                                   *
' *                                                                   *
' * ClassName: this shared, read-only property returns the cool name  *
' *      of this class: quickBasicEngine                              *
' *                                                                   *
' * clear: this method resets the engine to a start state by ensuring *
' *      all reference variables are cleared. You don't have to       *
' *      execute it in the normal case as long as you use dispose to  *
' *      responsibly clean up the compiler: dispose will clear.       *
' *                                                                   *
' * clearStorage: this method clears all variables to their default   *
' *      values appropriate to their type.                            *
' *                                                                   *
' * clone: this method implements ICloneable and returns a clone of   *
' *      the instance object. The clone is guaranteed to have the same*
' *      global properties such as SourceCode but it won't necessarily*
' *      be in the identical state as the instance object: the clone  *
' *      will be in the unscanned, uncompiled state.                  * 
' *                                                                   *
' *      The clone when passed to the compareTo method will return    *
' *      True.                                                        *
' *                                                                   *
' *      The clone method implements the ICloneable interface.        *
' *                                                                   *
' * codeChangeEvent(s,op,opIndex): this event is Raised when an op    *
' *      code is changed. It supplies the handle of the sender quick- *
' *      BasicEngine, the modified opcode as an object of type        *
' *      qbPolish, and the index of the modified opcode.              *
' *                                                                   *
' * codeGenEvent(s,op): this event is Raised when an op code is       *
' *      generated, and it supplies the handle of the sender quick-   *
' *      BasicEngine and the op as an object of type qbPolish.        *
' *                                                                   *
' * codeRemoveEvent(s,opIndex): this event is Raised when the Polish  *
' *      op code indexed by opIndex is removed; this typically happens*
' *      inside the assembler which removes remarks and labels. s is  *
' *      the handle of the quickBasicEngine sending the event; the    *
' *      opIndex is the index of the opcode from 1.                   *
' *                                                                   *
' * codeType(code): this Shared method returns the type of "code" as a*
' *      string:                                                      *
' *                                                                   *
' *      *  immediateCommand: code is a valid expression              *
' *      *  program: code is a valid executable program               *
' *      *  invalid: code is not valid                                *
' *                                                                   *
' * compareTo(o): this method compares the instance object to the     *
' *      quickBasicEngine o and returns True when o clones the        *
' *      instance. o clones the instance when the source code of o and*
' *      that of the instance is identical except for white space and *
' *      all global options such as the ConstantFolding property are  *
' *      the same.                                                    * 
' *                                                                   *
' *      The compareTo method implements the IComparable interface.   *
' *                                                                   *
' * compile: this method compiles the source code without interpreting*
' *      the source; if the source has not been scanned this method   *
' *      scans the source in full.                                    *
' *                                                                   *
' * Compiled: this read-only property returns True if the current     *
' *      source code has been compiled, False otherwise               *
' *                                                                   *
' * compileErrorEvent(s,m,i,CL,L, h,c): this event is raised when a   *
' *      user error is detected in the parser. In this event:         *
' *                                                                   *
' *      *  s is the handle of the sender quickBasicEngine            *
' *      *  m is the error message                                    *
' *      *  i is the character index at which the error was found     *
' *      *  CL is the length of the relevant surrounding context      *
' *      *  L is the line number                                      *
' *      *  h is additional help information                          *
' *      *  c is the error line of code                               *
' *                                                                   *
' * ConstantFolding: this read-write property returns and may be set  *
' *      to True or False:                                            *
' *                                                                   *
' *      *  When ConstantFolding is True then all subexpressions in   *
' *         the source code that consist exclusively of constants and *
' *         operators are evaluated by the compiler and not at run    *
' *         time. This can speed up run time.                         * 
' *                                                                   *
' *         For example, in a+1+1 when ConstantFolding is True, the   *
' *         subexpression 1+1 is evaluated by the compiler. This      *
' *         example is contrived (as in stupid) but many code and     *
' *         business rule generators may generate such examples.      * 
' *                                                                   *
' *      *  When ConstantFolding is False then all subexpressions in  *
' *         the source code that consist exclusively of constants and *
' *         operators are compiled normally to code.                  * 
' *                                                                   *
' * DegenerateOpRemoval: this read-write property returns and may be  *
' *      set to True or False:                                        *
' *                                                                   *
' *      *  When DegenerateOpRemoval is True then all operations      *
' *         known to have no effect at compile time (including        *
' *         addition of the constant zero and multiplication by the   *
' *         constant 1) are removed.                                  * 
' *                                                                   *
' *         For example, in a+0 when DegenerateOpRemoval is True, the *
' *         bogus addition is eliminated, and the only code generated *
' *         pushes the value of a on the stack. This example is       *
' *         contrived (as in stupid) but many code and business rule  *
' *         generators may generate such examples.                    * 
' *                                                                   *
' *         By calling these operations "degenerate" I do not mean    *
' *         anything naughty but instead use a mathematical term of   *
' *         art.                                                      *
' *                                                                   *
' *      *  When DegenerateOpRemoval is False then all degenerate     *
' *         operations generate code for run-time evaluation.         * 
' *                                                                   *
' * dispose: this method disposes of the object, and it cleans up     *
' *      reference objects in the heap.  For best results use this    *
' *      method when you are finished using the object in code. This  *
' *      method will mark the object as unusable.                     *
' *                                                                   *
' *      By default the compiler is inspected when it is disposed; to *
' *      suppress this check use the overlay dispose(False).          *
' *                                                                   *
' * EasterEgg: this Shared and read-only property returns a string    *
' *      containing all the About information and a dedicatory        *
' *      Easter Egg in the form of text only. There is no nonsense    *
' *      emanating from this Easter Egg in the form of images, sounds *
' *      or computer viruses, just plain Ascii text: honi soit qui    *
' *      mal y pense.                                                 *
' *                                                                   *
' * eval(s): this Shared method evaluates the string s as a single    *
' *      expression in Quick Basic notation, or as a series of        *
' *      statements (separated by colons), and followed by an expres- *
' *      sion. For example, s may be a series of Let assignment       *
' *      statements which set variable values, followed by an         *
' *      expression.                                                  *
' *                                                                   *
' *      The value of the expression is returned, as a qbVariable     *
' *      object.                                                      *
' *                                                                   *
' *      The eval method is "lightweight": see also evaluate. The eval*
' *      method creates a New quickBasicEngine having all default     *
' *      values and default properties, and the evaluated string is   *
' *      executed using default values and default properties.        *
' *                                                                   *
' *      The overload eval(s,ss), where ss is a reference string,     *
' *      places the event log produced in the evaluation in ss.       *
' *                                                                   *
' *      By default, the eval method returns the qbVariable which is  *
' *      the value of the expression; when the expression contains an *
' *      error, eval returns Nothing. The overload eval(s,flag)       *
' *      (where flag is a Boolean flag passed by reference) returns   *
' *      the string value of the expression or an error report        *
' *      when the expression contains an error. The flag is set to    *
' *      True on no error and False on an error.                      *
' *                                                                   *
' * evaluate(s): this method evaluates the string s as a single       *
' *      expression in Quick Basic notation, or as a series of        *
' *      statements (separated by colons), followed by a colon and a  *
' *      final expression. For example, s may be a                    *
' *      series of Let assignment statements which set variable       *
' *      values, followed by an expression.                           *
' *                                                                   *
' *      The value of the expression is returned, as a qbVariable     *
' *      object.                                                      *
' *                                                                   *
' *      The evaluate method is "heavyweight": see also eval. The     *
' *      evaluate method is executed inside the existing              *
' *      quickBasicEngine, and using the current setting of all values*
' *      and properties.                                              *
' *                                                                   *
' *      The overload evaluate(s,ss), where ss is a reference string, *
' *      places the event log produced in the evaluation in ss.       *
' *                                                                   *
' *      The overload evaluate(s,o), where o is an object, places the *
' *      .Net value of the evaluate in o.                             *
' *                                                                   *
' * Evaluation: this read-only property returns the result of the most*
' *      evaluate method. It returns Nothing when no such result      *
' *      exists.                                                      *
' *                                                                   *
' *      The result of the most recent evaluation is returned, as a   *
' *      qbVariable object.                                           *
' *                                                                   *
' * evaluationValue: this method returns the result of the most recent*
' *      evaluate method. It returns Nothing when no such             *
' *      result exists.                                               *
' *                                                                   *
' *      The result of the most recent evaluation is returned, as a   *
' *      .Net object.                                                 *
' *                                                                   *
' * EventLog: this read-only property returns a Collection of items,  *
' *      each of which represents an event raised by a RaiseEvent in  *
' *      the Quick Basic engine.                                      *
' *                                                                   *
' *      The event log is populated only then the EventLogging        *
' *      parameter has been set to True.                              *
' *                                                                   *
' *      This property is useful in finding events issued by a        *
' *      Quick Basic engine that has not been defined using           *
' *      WithEvents.                                                  *
' *                                                                   *
' *      Each item in the event log is a three-item subcollection:    *
' *      item(1) identifies the event: item(2) is the event date and  *
' *      time, and item(3) is a list of the event operands. In item(3)*
' *      each operand is separated by a new line and is in the format *
' *      name=value.                                                  *
' *                                                                   *
' * eventLog2ErrorList: this method produces a list of compiler-detec-*
' *      ted errors in source code and interpreter-detected errors in *
' *      logic, and it has the following overloads:                   *
' *                                                                   *
' *           *  eventLog2ErrorList with no operands returns the list *
' *              using the event log of the Quick Basic engine        *
' *              instance.                                            *
' *                                                                   *
' *           *  The Shared overload eventLog2ErrorList(c) returns the*
' *              list based on the collection c, which must be in the *
' *              format returned by the EventLog property of a        *
' *              Quick Basic engine.                                  *
' *                                                                   *
' *      The list is a sequence of newline-separated lines. Each line *
' *      will consist of the error type (compilerError or             *
' *      interpreterError), followed by a colon, a space, and the text*
' *      of the error.                                                *
' *                                                                   *
' * eventLogFormat: this method formats the event log and returns     *
' *      the formatted log. The log is best shown using a monospaced  *
' *      font such as Courier New.                                    *                                                
' *                                                                   *
' *      This method has three overloads.                             *
' *                                                                   *
' *           *  eventLogFormat: returns the entire log               *
' *                                                                   *
' *           *  eventLogFormat(n): returns the log starting at entry *
' *              n                                                    *
' *                                                                   *
' *           *  eventLogFormat(n,c): returns c entries of the log    *
' *              starting at entry n                                  *
' *                                                                   *
' * EventLogging: this read-write property returns and may be set to  *
' *      True or False to control and determine the population of the *
' *      EventLog as described above.                                 *
' *                                                                   *
' * ExtensionAvailable: this Shared, read-only property returns True  *
' *      when the compile-time symbol QUICKBASICENGINE_EXTENSION is   *
' *      True, False otherwise.                                       *
' *                                                                   *
' * GenerateNOPs: this read-write property returns and may be set to  *
' *      True when the compiler should generate NOP and REM instruc-  *
' *      tions as part of the object code, False otherwise.           *
' *                                                                   *
' *      This property is exposed only when the QUICKBASICENGINE_     *
' *      EXTENSION compile-time flag has been set to True. Its default*
' *      value is True.                                               *
' *                                                                   *
' * getThreadStatus: this method returns the thread status while      *
' *      DISCOUNTING its own effect as one of the strings             *
' *      Initializing, Ready, Running, Stopping or Stopped.           *
' *                                                                   *
' *      See also runningThreads.                                     *
' *                                                                   *
' *      Note: the value returned by getThreadStatus is               *
' *      nondeterministic because while getThreadStatus discounts its *
' *      own effect, the status may change while getThreadStatus is   *
' *      executing.                                                   *
' *                                                                   *
' *      For this reason, getThreadStatus should be used for adult    *
' *      entertainment purposes only, for example to display the      *
' *      nondeterministic status in a GUI.                            *
' *                                                                   *
' * inspect(r): this method inspects the object. It checks for errors *
' *      that result from my stupid blunders in the creating the      *
' *      original source code of this class, your ham-fisted changes  *
' *      to the source code of this class, or "object abuse", the use *
' *      of this object after an error has occured. Inspect does not  *
' *      look for simple user errors: these are prevented elsewhere.  *
' *                                                                   *
' *      r should be a string, passed by reference; it is assigned an *
' *      inspection report. The r parameter may be omitted.           *
' *                                                                   *
' *      The following inspection rules are used.                     *
' *                                                                   *
' *           *  The object instance must be usable                   *
' *                                                                   *
' *           *  The scanner object must pass its own inspection      *
' *                                                                   *
' *           *  The collection of qbPolish instructions must contain *
' *              qbPolish objects exclusively. If the collection      *
' *              contains fewer than 101 objects then each object must*
' *              pass the qbPolish.inspect inspection; if more than   *
' *              100 objects exist then a random selection of objects *
' *              is inspected.                                        * 
' *                                                                   *
' *           *  The collection of qbVariable variables must conform  *
' *              to the structure described in the preceding section. * 
' *                                                                   *
' *           *  The subroutine and function index must conform       *
' *              to the structure described in the preceding section. * 
' *                                                                   *
' *           *  The constant expression index must conform           *
' *              to the structure described in the preceding section. * 
' *                                                                   *
' *      If the inspection is failed the object becomes unusable.     *
' *                                                                   *
' *      An internal inspection is carried out in the constructor and *
' *      inside the dispose method.                                   *
' *                                                                   *
' * InspectCompilerObjects: this read-write property returns and may  *
' *      be set to True when objects created by the compiler need to  *
' *      be inspected when disposed.                                  *
' *                                                                   *
' *      Set this option to True when testing the compiler and modifi-*
' *      cations as a way to be sure that objects don't include       *
' *      buggy code. Its default is False, and note that setting this *
' *      option will slow the compiler down.                          *
' *                                                                   *
' *      When this option is True the following compiler object types *
' *      will be inspected when they are disposed.                    *
' *                                                                   *
' *                                                                   *
' *      *  The scanner                                               *
' *                                                                   *
' *      *  Each variable that is created during compilation and      *
' *         interpretation (including its type)                       *
' *                                                                   *
' *      *  The Quick Basic engine                                    *
' *                                                                   *
' * interpret: this method interprets the compiled code (it will scan,*
' *      compile and assemble the source code as needed.) This method *
' *      will return an Object:                                       * 
' *                                                                   *
' *                                                                   *
' *      *  If the stack is empty at the end of interpretation, this  *
' *         method will return True                                   *
' *                                                                   *
' *      *  If the stack contains one entry at the end of interpreta- *
' *         tion this method returns that entry, which will be a      *
' *         qbVariable                                                *
' *                                                                   *
' *      *  If the stack contains multiple entries at the end of      *
' *         interpretation, this method will return False             *
' *                                                                   *
' *                                                                   *
' *      Note that the interpret method does Quick Basic input and    *
' *      output by means of events: see the interpretInputEvent and   *
' *      the interpretPrintEvent for details.                         *
' *                                                                   *
' *      By default, the interpret method will inspect temporary      *
' *      stack variables when they are disposed.                      *
' *                                                                   *
' * interpretClsEvent(s): this event is Raised when the CLS           *
' *      instruction is executed to clear the output screen. It       *
' *      supplies the handle of the sender quickBasicEngine.          *
' *                                                                   *
' * interpretErrorEvent(s,m,i,h): this event is raised for an error   *
' *      detected by the interpreter that is probably a logic error   *
' *      in the Quick Basic code. In this event:                      *
' *                                                                   *
' *      *  s is the handle of the sender quickBasicEngine            *
' *      *  m is the error message                                    *
' *      *  i is the index of the failing Polish instruction          *
' *      *  h is additional help text                                 *
' *                                                                   *
' * interpretInputEvent(s,c): this event is raised when the interpret-*
' *      er needs data from the virtual quickBasic console            *
' *      maintained by the GUI. s is the handle of the sender quickBa-*
' *      sicEngine and c is a string passed bt reference. Its handler *
' *      should get the input and set c to the input string.          *
' *                                                                   *
' * interpretPrintEvent(s,str): this event is raised to print or      *
' *      otherwise display the output string str. s is the handle of  *
' *      the quickBasicEngine sender and str is the output string.    *
' *                                                                   *
' * interpretTraceEvent(s,i,stack,storage): this event is raised each *
' *      time the interpreter executes a Polish instruction. In this  *
' *      event, s is the handle of the sender quickBasicEngine, i is  * 
' *      the Polish instruction stack as an object of                 *
' *      type stack, and storage is the collection of variables (each *
' *      item is of type qbVariable).                                 *
' *                                                                   *
' *      To obtain this event, the quickBasicEngine must be compiled  *
' *      with its compile-time symbol QUICKBASICENGINE_POPCHECK set to*
' *      True: see also interpretTraceFastEvent.                      *
' *                                                                   *
' *      Note that the trace event occurs BEFORE the operation is     *
' *      executed.                                                    *
' *                                                                   *
' *      The Shared method stack2String is available to serialize     *
' *      the stack: the Shared method storage2String is available to  *
' *      serialize the storage.                                       *
' *                                                                   *
' * interpretTraceFastEvent(s,i,stack,storage): this event is raised  *
' *      each time the interpreter executes a Polish instruction. In  *
' *      this event, s is the handle of the sender quickBasicEngine, i* 
' *      is the Polish instruction stack as an array of type          *
' *      qbVariable, and storage is the collection of variables (each *
' *      item is of type qbVariable).                                 *
' *                                                                   *
' *      To obtain this event, the quickBasicEngine must be compiled  *
' *      with its compile-time symbol QUICKBASICENGINE_POPCHECK set to*
' *      False: see also interpretTraceEvent.                         *
' *                                                                   *
' *      Note that the trace event occurs BEFORE the operation is     *
' *      executed.                                                    *
' *                                                                   *
' *      The Shared method stack2StringFast is available to serialize *
' *      the stack: the Shared method storage2String is available to  *
' *      serialize the storage.                                       *
' *                                                                   *
' *      Note: this event is not added to the event log.              *
' *                                                                   *
' * loopEvent(s,a,e,n,c,L,x): this event is raised for a number of Do *
' *      and For loops inside the quickBasicEngine and it shows       *
' *      progress.                                                    *
' *                                                                   *
' *      The loopEvent exposes the following parameters:              *
' *                                                                   *
' *      *  s: the handle of the sender quickBasicEngine.             *
' *                                                                   *
' *      *  a: a By Value string, which contains a description of the *
' *         activity in the Do or For loop.                           *
' *                                                                   *
' *      *  e: a By Value string, which contains a description of the *
' *         entity being processed in the Do or For loop.             *
' *                                                                   *
' *      *  n: a By Value integer, which contains the current index,  *
' *         from 1, of the entity being processed in the Do or For.   *
' *                                                                   *
' *      *  c: a By Value integer, which contains the total number of *
' *         entities being processed in the loop.                     *
' *                                                                   *
' *      *  L: a By Value integer, which contains the nesting level of*
' *         the Do or For within other loops.                         *
' *                                                                   *
' *      *  x: extended comments                                      *
' *                                                                   *
' * mkUnusable: this method makes the object not usable.              *
' *                                                                   *
' * msgEvent(s,m): this event sends the text message m to its handler.*
' *      s is the handle of the sender quickBasicEngine.              *
' *                                                                   *
' * msilRun: this method translates the intepreted code into Microsoft*
' *      Intermediate Language, generates a simple dynamic assembly,  *
' *      and runs the code. This method returns the final stack value *
' *      as a function value.                                         *
' *                                                                   *
' *      At this writing only a few Polish opcodes are translatable to*
' *      MSIL. If Polish operations exist that cannot be translated,  *
' *      this method throws an error and returns Nothing. If the MSIL *
' *      code leaves an empty stack this method throws an error and   *
' *      returns Nothing.                                             *
' *                                                                   *
' * Name: this read-write property returns and can change the name of *
' *      the object instance. Name defaults to quickBasicEnginennnn   *
' *      date time, where nnnn is an object sequence number.          *
' *                                                                   *
' * object2XML: this method converts the state of the object to an    *
' *      eXtendedMarkupLanguage string; note that the returned tag    *
' *      will include all source code and all parsed tokens and as    *
' *      such may be unmanageably large for large source code files.  *
' *                                                                   *
' *      These optional parameters are exposed:                       *
' *                                                                   *
' *           *  booAboutComment:=False will suppress a boxed XML     *
' *              comment at the start of the XML containing the value *
' *              of the About property of this object.                *
' *                                                                   *
' *           *  booStateComment:=False will suppress comments that   *
' *              describe each state value returned.                  *
' *                                                                   *
' * PopCheck: this Shared, read-only property returns True when the   *
' *      compile-time symbol QUICKBASICENGINE_POPCHECK is True,      *
' *      False otherwise.                                             *
' *                                                                   *
' * parseEvent(s,gc,t,si,L,tsi,tl,ps,pL,c,v): this event fires at the *
' *      completion of each successful parse of any terminal or       *
' *      nonterminal grammar category and it is useful for progress   *
' *      reporting.                                                   * 
' *                                                                   *
' *      In this event:                                               * 
' *                                                                   *
' *      *  s is the handle of the sender quickBasicEngine            *
' *                                                                   *
' *      *  gc is the name of the grammar category                    *
' *                                                                   *
' *      *  t is True when the grammar category is a terminal         *
' *                                                                   *
' *      *  si is the CHARACTER index in the source code at which the *
' *         category starts                                           *
' *                                                                   *
' *      *  L is the CHARACTER length of the source code that corres- *
' *         ponds to the category                                     *
' *                                                                   *
' *      *  tsi is the TOKEN index in the source code at which the    *
' *         category starts                                           *
' *                                                                   *
' *      *  tl is the TOKEN length of the source code that corres-    *
' *         ponds to the category                                     *
' *                                                                   *
' *      *  ps is the index in the object code at which the compiled  *
' *         category starts                                           *
' *                                                                   *
' *      *  pL is the length in tokens of the code for the category   *
' *                                                                   *
' *      *  c is a comment associated with this parse                 *
' *                                                                   *
' *      *  v is the compilation level                                *
' *                                                                   *
' * parseFailEvent(s,gc): this event fires at the failure of the parse*
' *      event for the grammar category identified by gc. s is the    *
' *      handle of the sender quickBasicEngine.                       *
' *                                                                   *
' * parseStartEvent(s,gc): this event fires at the start of the parse *
' *      event for the grammar category identified by gc. s is the    *
' *      handle of the sender quickBasicEngine.                       *
' *                                                                   *
' * PolishCollection: this read-only property returns the collection  *
' *      of Polish operations compiled from the source code. Note that*
' *      if no code has been compiled this collection will be Nothing.*
' *                                                                   *
' * PopCheckAvailable: This Shared, read-only property returns True   *
' *      when the quickBasicEngine exposing it was generated with the *
' *      compile-time symbol QUICKBASICENGINE_POPCHECK set to True,   *
' *      False otherwise.                                             *
' *                                                                   *
' * reset: this method resets the Quick Basic engine                  *
' *                                                                   *
' * resumeQBE: this method puts the Quick Basic Engine in the Ready   *
' *      state when it is in the Stopped state. If the object is in   *
' *      any other state, resume has no effect and results in no      *
' *      error.                                                       *
' *                                                                   *
' *      For best results clear the QBE after resuming it.            *
' *                                                                   *
' * run: this method runs the immediate command or source program     *
' *      that is the value of the SourceCode property. As needed, this*
' *      method will scan, parse, and/or assemble the code.           *
' *                                                                   *
' *      The optional overload run(s) will run the code as an         *
' *      immediate command when s is immediateCommand, or as a        *
' *      program when s is program. If s is not supplied the command  *
' *      type is determined, by first executing the code as an        *
' *      immediate command and then on failure executing the code as  *
' *      a program.                                                   *
' *                                                                   *
' * runningThreads: this method returns the number of threads that    *
' *      are running procedures inside the Quick Basic engine as a    *
' *      number between 0 and n. This method INCLUDES its own thread. *
' *                                                                   *
' *      See also getThreadStatus.                                    *
' *                                                                   *
' *      Note: the value returned by runningThreads is                *
' *      nondeterministic because the status may change while         *
' *      runningThreads is executing. The value will always be one    *
' *      or greater because runningThreads INCLUDES its own thread.   *
' *                                                                   *
' *      For this reason, runningThreads should be used for adult     *
' *      entertainment purposes only, for example to display the      *
' *      nondeterministic count in a GUI.                             *
' *                                                                   *
' * scan: this method scans the source code.                          *
' *                                                                   *
' * scanEvent(s,t): this event is raised when the scanner scans a     *
' *      token on behalf of the compiler. s is the handle of the send-*
' *      quickBasicEngine; t is a qbToken object. To Handle this      *
' *      event, your code needs to reference qbToken and qbTokenType. *
' *                                                                   *
' *      Note that the scanEvent identifies only the scanned token;   *
' *      see the loopEvent, which is also raised when tokens are      *
' *      scanned. The loopEvent identifies the scan character position*
' *      and the total number of characters to scan.                  *
' *                                                                   *
' * Scanned: this read-only property returns True if the current      *
' *      source code has been scanned, otherwise False.               *
' *                                                                   *
' * Scanner: this read-only property returns the qbScanner object,    *
' *      used to scan for parsing. The parse has-a qbScanner.         *
' *                                                                   *
' * SourceCode: this read-write property returns and may be set to    *
' *      the source code for parsing. Assigning source code clears    *
' *      the array of tokens in the object state, but does not result *
' *      in an immediate scan of the source code. Instead, scanning   *
' *      occurs when the QBToken property is called, and the token is *
' *      not available.                                               *
' *                                                                   *
' * stack2String(stack): this Shared method serializes the interpreter*
' *      stack, which is returned by the InterpretTraceEvent.         *
' *                                                                   *
' *      This method is generated when the compile-time symbol        *
' *      QUICKBASICENGINE_POPCHECK is True.                           *
' *                                                                   *
' *      The serialized stack consists of the stack entries (starting *
' *      at the top of the stack). Each entry is a qbVariable that    *
' *      has been converted to a string using qbVariable.toString.    *
' *      Entries are divided by newline characters.                   *
' *                                                                   *
' * stack2StringFast(stack, stackTop): this Shared method serializes  *
' *      the interpreter stack, which is returned by the              *
' *      InterpretTraceFastEvent.                                     *
' *                                                                   *
' *      This method is generated when the compile-time symbol        *
' *      QUICKBASICENGINE_POPCHECK is False.                          *
' *                                                                   *
' *      The stack parameter should be the stack as an array of       *
' *      qbVariable objects. The stackTop should be the index of the  *
' *      top of the stack.                                            *
' *                                                                   *
' *      The serialized stack consists of the stack entries (starting *
' *      at the top of the stack). Each entry is a qbVariable that    *
' *      has been converted to a string using qbVariable.toString.    *
' *      Entries are divided by newline characters.                   *
' *                                                                   *
' * stopQBE: this method puts the Quick Basic Engine in the Stopped   *
' *      state when it is in the Ready or Running state. If the object*
' *      is in the Stopped state already, stopQBE has no effect and   *
' *      results in no error.                                         *
' *                                                                   *
' *      If the object is in the Running state, it is "quiesced" as   *
' *      follows.                                                     *
' *                                                                   *
' *      *  Any executing For or Do loop in the engine, which issues  *
' *         loop events, is exited as soon as the loop event is       *
' *         issued.                                                   *
' *                                                                   *
' *      *  If the parser is running the compiler is exited           *
' *         when the next grammar category is recognized.             *
' *                                                                   *
' *      *  If the interpreter is running the interpreter is exited   *
' *         as soon as the interpreter's loop event is issued.        *
' *                                                                   *
' * storage2String(stack): this Shared method serializes the          *
' *      interpreter storage, which is returned by the                *
' *      InterpreterTraceEvent.                                       *
' *                                                                   *
' *      The serialized storage consists of the stack entries         *
' *      (starting at the top of the stack). Each entry is a          *
' *      qbVariable that has been converted to a string using         *
' *      qbVariable.toString. Entries are divided by newline          *
' *      characters.                                                  *
' *                                                                   *
' * Tag: this read-write property returns and may be set to an object *
' *      or value to be associated with the Quick Basic engine.       *
' *                                                                   *
' * test(r): this method creates a test engine and then runs a        *
' *      series of tests on the real created engine.                  *
' *                                                                   *
' *      r should be a string, passed by reference. It is set to a    *
' *      test report.                                                 *
' *                                                                   *
' *      If all tests are passed this method returns True. If any test*
' *      fails, this method returns False and marks the object not    *
' *      usable.                                                      *
' *                                                                   *
' *      At this writing the tests consist  of a series of evaluated  *
' *      expressions and small programs with their expected results.  *
' *                                                                   *
' *      Note that the overload test(r, True) will include a rather   *
' *      large event trace (using the EventLogging feature of the     *
' *      quickBasicEngine) in the test report.                        *
' *                                                                   *
' *      Note also that the test engine has default properties with   *
' *      three exceptions:                                            *
' *                                                                   *
' *           *  If constant folding optimization is in effect in the *
' *              main engine it will be in effect in the test engine. *
' *                                                                   *
' *           *  If degenerate operation removal is in effect in the  *
' *              main engine it will be in effect in the test engine. *
' *                                                                   *
' *           *  If removal of labels and comments in assembly is in  *
' *              effect in the main engine it will be in effect in the*
' *              test engine.                                         *
' *                                                                   *
' * threadStatusChangeEvent(s): this event is raised with no operands *
' *      when the number of threads running qbe code changes or the   *
' *      qbe is stopped. s is the handle of the sender                *
' *      quickBasicEngine                                             *
' *                                                                   *
' * toString: this override method returns the Name property of the   *
' *      quickBasicEngine.                                            *
' *                                                                   *
' * Usable: this read-only property returns True if the object is     *
' *      usable, False otherwise.                                     *
' *                                                                   *
' * userErrorEvent(s,d,h): this event occurs when there has been a    *
' *      stupid error we can blame on the user of the quickBasicEngine*
' *      considered as a Quick Basic system and not as a .Net object, *
' *      as opposed to my ham-fisted blunders in the actual code of   *
' *      this object or idiotic mistakes in the graphical user        *
' *      interface.                                                   *
' *                                                                   *
' *      s is the handle of the sender quickBasicEngine: d is the     *
' *      error description and h is additional help.                  *
' *                                                                   *
' * Variable(i): this indexed, read-write property returns and can    *
' *      be set to qbVariable objects. The index i may be the variable*
' *      name or it may be the variable index from 1.                 *
' *                                                                   *
' *      When this property is assigned, the assigned variable must   *
' *      have the same name as the variable it replaces.              *
' *                                                                   *
' * VariableCount: this read-only property returns the current number *
' *      of variables found in the source code.                       *
' *                                                                   *
' *                                                                   *
' * MULTITHREADING CONSIDERATIONS ----------------------------------- *
' *                                                                   *
' * Multiple, distinct instances of this object may run simultaneously*
' * in multiple threads; the same object instance may run simultane-  *
' * ous procedures in multiple threads, although locking may occur:   *
' * this is because this full-thread capability is implemented by     *
' * locking the object on entry to each Public and Friend procedure   *
' * except for selected methods that directly check or change the     *
' * running thread count, and procedure.                              *
' *                                                                   *
' * The standard policy, on entry to a Public procedure that does not *
' * check or modify thread counts or immediately call a common pro-   *
' * cedure, is to call a dispatch_ routine with the following         *
' * arguments:                                                        *
' *                                                                   *
' *                                                                   *
' *      *  The name of the procedure as a string, including Get or   *
' *         Set when the procedure is a property                      *
' *                                                                   *
' *      *  The default value to be returned when the Quick Basic     *
' *         engine is not available, or the qbe is unusable.          *
' *                                                                   *
' *      *  The message to be shown when the Quick Basic engine is    *
' *         not available or not usable                               *
' *                                                                   *
' *      *  Zero, one or more parameters to be passed to the          *
' *         procedure                                                 *
' *                                                                   *
' *                                                                   *
' * The following is the shape of the dispatch_ routine:              *
' *                                                                   *
' *                                                                   *
' *        SyncLock OBJthreadStatus                                   *
' *            If Not checkAvailable_(name,help) Then Return...       *
' *            OBJthreadStatus.startThread                            * 
' *        End SyncLock                                               *
' *        SyncLock OBJstate                                          *
' *            With OBJstate.usrState                                 *
' *                If Not checkUsable_(name,help) Then Return...      *
' *                Select Case UCase(strProcedure)                    *
' *                .                                                  *
' *                .                                                  *
' *                .                                                  *
' *                End Select                                         *
' *            End With                                               *
' *        End SyncLock                                               *
' *        SyncLock OBJthreadStatus                                   *
' *            OBJthreadStatus.stopThread()                           * 
' *        End SyncLock                                               *
' *                                                                   *
' *                                                                   *
' * When adding any Public procedure, pass it to dispatch_ as shown   *
' * above UNLESS it needs to interact (like the existing methods      *
' * getThreadStatus, resumeQBE, runningThreads and stopQBE) directly  *
' * with the thread status. In addition, and unless the new procedure *
' * interacts with thread status, a case that handles the new         *
' * procedure has to be added to the Select statement of dispatch_.   *
' *                                                                   *
' * This case should carry out the goals of the new procedure, and    *
' * in most cases assign the function method's or Get procedure's     *
' * return value to objReturn, which is passed back to the caller.    *
' *                                                                   *
' * Note that this object is designed for low-intensity multithreading*
' * because it locks the state in a rather broad fashion on entry to  *
' * each public procedure and because it uses objects which are not   *
' * thread safe. At this writing the intent of making the Quick Basic *
' * Engine threadsafe is to allow it to be stopped in a GUI. Using the*
' * same instance to run multiple threads, other than quick threads to*
' * change and determine the run state, will result in locking, lower *
' * performance, and erroneous results at this time.                  *
' *                                                                   *
' * Whether or not threads are used, the Quick Basic engine is in one *
' * of these states:                                                  *
' *                                                                   *
' *                                                                   *
' *      *  Ready to run                                              *
' *      *  Running in one or more threads                            *
' *      *  Stopping                                                  *
' *      *  Stopped                                                   *
' *      *  Unusable (run status is moot)                             *
' *                                                                   *
' *                                                                   *
' * The stopQBE method will stop the Quick Basic engine. The mkReady  *
' * method will resume a stopped Quick Basic engine to Ready status   *
' * (no provision is made for resuming a stopped run: when the engine *
' * is stopped results are lost).                                     *
' *                                                                   *
' *                                                                   *
' * COMPILE-TIME SYMBOLS -------------------------------------------- *
' *                                                                   *
' * The compile-time symbol QUICKBASICENGINE_EXTENSION should         *
' * be set to True to generate support for the following extensions:  *
' *                                                                   *
' *                                                                   *
' *      *  Quick Basic strings that are not limited to 64K bytes.    *
' *                                                                   *
' *         By default, Quick Basic strings as implemented by this    *
' *         compiler are restricted to 64K bytes essentially as a trip*
' *         down Memory Lane and in honor of programmers who developed*
' *         workarounds for these limits in the past.                 *
' *                                                                   *
' *         However, you may compile this source code with the compile*
' *         time symbol QUICKBASICENGINE_EXTENSION set to False to    *
' *         support strings limited only in the same way .Net strings *
' *         are limited.                                              *
' *                                                                   *
' *      *  The GenerateNOPs property. This property, when it has its *
' *         default value of True, causes the compiler to generate    *
' *         explanatory REM and commented NOP instructions when creat-*
' *         ing the object interpreter code.                          *
' *                                                                   *
' *         When this property is set to False, REMs and NOPs are not *
' *         generated.                                                *
' *                                                                   *
' *                                                                   *
' * The compile-time symbol QUICKBASICENGINE_POPCHECK should be set   *
' * to True or False:                                                 *
' *                                                                   *
' *                                                                   *
' *      *  When this symbol is True, a version of the interpreter    *
' *         will be generated that runs rather slowly, making extra   *
' *         checks on the validity of the stack.                      *
' *                                                                   *
' *      *  When this symbol is False, a version of the interpreter   *
' *         will be generated that runs faster but which doesn't      *
' *         check stack entries.                                      *
' *                                                                   *
' *                                                                   *
' * ERROR TAXONOMY -------------------------------------------------- *
' *                                                                   *
' * Three types of error potentials are recognized by this object.    *         
' *                                                                   *
' *                                                                   *
' *      *  Bonehead errors in logic, which are of two subtypes.      *
' *                                                                   *
' *         + My stupid blunders remaining after my thorough testing  *
' *         + Your idiotic errors in making changes to this code      *
' *                                                                   *
' *      *  Ham-fisted mistakes made in using the object interface    *
' *         whether this is a GUI or an object using the Quick Basic  *
' *         engine as a service.                                      *
' *                                                                   *
' *      *  Moronic errors in Quick Basic coding and logic            *
' *                                                                   *
' *                                                                   *
' * These errors are handled in three ways:                           *
' *                                                                   *
' *                                                                   *
' *      *  Errors in the logic of the code, that are detected by this*
' *         code itself (such as in the inspect method) will result   *
' *         in calls to the low level errorHandler utility exposed by *
' *         the utilities object, which typically displays a MsgBox.  *
' *                                                                   *
' *         These errors will usually mark the instance object as     *
' *         not usable so it does not damage your data.               *
' *                                                                   *
' *      *  Mistakes in calling the object also result in calls to    *
' *         utilities.errorHandler but do not mark the object as      *
' *         unusable.                                                 * 
' *                                                                   *
' *      *  Errors in coding Quick Basic and errors in logic are      *
' *         handled by events: see the section "Properties, Methods   *
' *         and Events" for more information.                         *
' *                                                                   *
' *                                                                   *
' * REFERENCES ------------------------------------------------------ *
' *                                                                   *
' * Hergert 1989: Douglas Hergert, Microsoft QuickBasic, Third Ed.    *
' * Redmond, WA: Microsoft Press                                      *
' *                                                                   *
' *********************************************************************

' *********************************************************************
' *                                                                   *
' * To Darlene, Eddie and Peter (junglee Peter): for in dreams begin  *
' * responsibilities                                                  *
' *                                                                   *
' * "Computer science is no more about computers than astronomy is    *
' *  about telescopes."                                               *
' *                                                                   *
' *                       - Edsger Dijkstra 1930..2002                *
' *                                                                   *
' * "But the man who knows the relation between the forces of nature  *
' *  and action, sees how some forces of Nature work upon other forces*
' *  of Nature, and becomes not their slave."                         *
' *                                                                   *
' *                       - Bhagavad-Gita                             *
' *                                                                   *
' * "I could be bound in a nutshell and count myself king of absolute *
' *  space."                                                          *
' *                                                                   *
' *                       - Shakespeare, Hamlet                       *
' *                                                                   *
' *********************************************************************

Public Class qbQuickBasicEngine

    Implements ICloneable
    Implements IComparable

    ' ***** Shared *****
    Private Shared _INTsequence As Integer
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJqbOp As qbOp.qbOp
    Private Shared _OBJqbVariable As qbVariable.qbVariable
    Private Shared _OBJqbVariableType As qbVariableType.qbVariableType

    ' ***** Tracking of nested structures *****
    Private Enum ENUnesting
        openCode                                    ' No control structure
        forLoop                                     ' For loop with control variable
        doLoop                                      ' New-style Do while|until with zero and one trip options
        whileLoop                                   ' Legacy while..wend (yecchhh)
        ifThen                                      ' Then side of an If statement
        elseClause                                  ' Else side of an If statement
    End Enum
    Private Structure TYPnesting
        Dim strStartLabel As String                 ' Label of For/Do/While loop start (null string for If)  
        Dim enuNestType As ENUnesting               ' For loop, structured do, legacy while, then clause, or else
        Dim objInfo As Object                       ' forLoop: control lValue as a string
                                                    ' Structured do: zero trips: 0
                                                    ' Structured do: one trip: 1
                                                    ' Legacy while: Nothing
                                                    ' Then clause: index of conditional goto instruction at start of then clause
                                                    ' Else clause: index of unconditional goto instruction at end of then clause
        Dim intLoopExits() As Integer               ' Do/while/for: zero, one or more jumps to code past end
        Dim intLine As Integer                      ' Line number
        Dim strStackImage As String                 ' Stack image, used at this time only for for
    End Structure

    ' ***** Subroutine and function table *****
    ' --- Formal parameter list
    Private Structure TYPformalParameters
        Dim strName As String               ' Parameter name
        Dim booByVal As Boolean             ' True: ByVal parameter: False: ByRef
    End Structure
    ' --- Subroutine/function structure
    Private Structure TYPsubFunction
        Dim booFunction As Boolean  ' True: Function
        Dim strName As String       ' Function and subroutine name
        Dim usrFormalParameters() As TYPformalParameters
        Dim intLocation As Integer  ' Address (0 for undef)
    End Structure

    ' ***** Constant Expressions *****
    Private Structure TYPconstantExpression
        Dim strConstantExpression As String
        Dim objValue As qbVariable.qbVariable
    End Structure

    ' ***** Tracing *****
    Private Structure TYPtracing
        Dim booTextTrace As Boolean                 ' True: provide the textual trace
        Dim booHeadsupTrace As Boolean              ' True: provide the heads-up trace
        Dim booTraceLines As Boolean                ' True: trace when source code line changes
                                                    ' False: trace every opcode
        Dim intRate As Integer                      ' Rate (usually 1 for every instruction or every line)
        Dim booIncludeSource As Boolean             ' True: include the source code
        Dim booIncludeMemory As Boolean             ' True: include memory
        Dim booIncludeStack As Boolean              ' True: include the stack
        Dim booIncludeObject As Boolean             ' True: include object code
        Dim booString2Box As Boolean                ' True: box the displays
    End Structure

    ' ***** Object state *****                 
    ' Checklist for adding a new variable to state                  
    '                                                               
    ' (1) Add it to the TYPstate structure right here.              
    '                                                               
    ' (2) If the variable represents a reference object then it     
    '     probably needs to be constructed in your New constructor. 
    '     If the variable represents a value object it may need to  
    '     be initialized.  
    '
    '     Some variables, however, may be constructed only when
    '     needed.
    '                                                               
    ' (3) Add it to the big statement in the object2XML method that 
    '     converts the state to eXtended Markup Language.           
    '                                                               
    '     If it is a value object, convert its value to a string    
    '     and convert the string to an XML tag as seen in prior art.
    '                                                               
    '     If it is a reference object, use its toString method or   
    '     the object2String utility to convert its value to a       
    '     string.                                                   
    '                                                               
    ' (4) If the variable has fewer valid values that it can contain
    '     or if the variable has invalid interactions with other     
    '     values in the state, then add tests for these invalid     
    '     values to the inspect method.                             
    '                                                               
    ' (5) If the variable represents a reference object with its    
    '     own inspect method you should add this subinspection to  
    '     the inspect method of this class.                         
    '                                                               
    ' (6) If the variable represents a reference object, you may need   
    '     at a minimum to set it to Nothing in the dispose method,  
    '     and ideally execute its dispose procedure. This won't be
    '     the case if the reference object is merely linked to your
    '     object like a "tag".
    '                                                               
    ' (7) If the variable represents a value that is set from outside
    '     the quickBasicEngine using a Public property, and is not part 
    '     of the internally developed quickBasicEngine state, then this 
    '     variable may have to be cloned (in the clone method) and/or 
    '     compared (in the compareTo method).
    '
    '     Such variables probably need also to be cleared in the clear
    '     method.
    '
    '     Existing examples include booConstantFolding.
    '
    ' (8) Add a description of the variable to the comment header
    '     section labeled "The quickBasicEngineState". 
    '                                                               
    ' 
    Private WithEvents OBJthreadStatus As threadStatus
    Private Structure TYPstate
        Dim booUsable As Boolean                                ' True: object is usable
        Dim strName As String                                   ' Object instance name
        Dim objScanner As qbScanner.qbScanner                   ' Scanner including source code
        Dim colPolish As Collection                             ' Type is qbPolish   
        Dim colVariables As Collection                          ' Variable collection  
        Dim objCollectionUtilities _
            As collectionUtilities.collectionUtilities          ' Utilities for collections 
        Dim booAssembled As Boolean                             ' True: code has been assembled
        Dim booCompiled As Boolean                              ' True: code has been compiled
        Dim enuSourceCodeType As enuSourceCodeType              ' Unknown, immediate or program
        Dim booConstantFolding As Boolean                       ' True: compiler evaluates constant expressions
        Dim booDegenerateOpRemoval As Boolean                   ' True: degenerate operations such as add in a+0 are removed
        Dim objImmediateResult As qbVariable.qbVariable         ' Nothing or eval/evaluate result
        Dim booExplicit As Boolean                              ' True: Option Explicit occurs
        Dim usrSubFunction() As TYPsubFunction                  ' Subroutine and function table             
        Dim colSubFunctionIndex As Collection                   ' Subroutine and function index
        Dim intLabelSeq As Integer                              ' Label sequence number
        Dim usrConstantExpression() As TYPconstantExpression    ' Constant expression table             
        Dim colConstantExpressionIndex As Collection            ' Constant expression index
        Dim intOldestConstantIndex As Integer                   ' Oldest constant expression
        Dim colReadData As Collection                           ' Data for the Read statement
        Dim intReadDataIndex As Integer                         ' Index for next Read
        Dim colLabel As Collection                              ' Key: stmt num or label preceded by _
                                                                ' Data: location
        Dim booAssemblyRemovesCode As Boolean                   ' True by default: remove REMs/labels
                                                                ' False: don't remove
        Dim colEventLog As Collection                           ' The event log (set to Nothing if
                                                                ' event logs are not in effect)
        Dim booInspectCompilerObjects As Boolean                ' True: inspect compiler objects
        Dim objTag As Object                                    ' Object tag
#If QUICKBASICENGINE_EXTENSION Then
        Dim booGenerateNOPs As Boolean                          ' True (default): generate NOPs
#End If
    End Structure
    Private Class stateClass
        Public usrState As TYPstate
    End Class
    Private OBJstate As stateClass
    ' Defined here but referenced in the state to be able to use WithEvents
    Private WithEvents OBJcollectionUtilities As collectionUtilities.collectionUtilities
    Private WithEvents OBJscanner As qbScanner.qbScanner

    ' ***** Test engine *****
    Private WithEvents OBJtestQBE As quickBasicEngine.qbQuickBasicEngine

    ' ***** Constants *****
    ' --- Easter egg
    Private Const ABOUT_INFO As String = _
        "This class compiles and interpretively runs a subset of the Quick Basic " & _
        "language. Note that the phrase Quick Basic is the intellectual property " & _
        "of the Microsoft corporation." & _
        vbNewLine & vbNewLine & _
        "This class was developed by " & _
        vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    Private Const EASTEREGG_INFO As String = _
        "To Darlene, Eddie and Peter (junglee Peter): for in dreams begin " & _
        "responsibilities." & _
        vbNewLine & vbNewLine & vbNewLine & _
        """Computer science is no more about computers than astronomy is " & _
        "about telescopes.""" & _
        vbNewLine & vbNewLine & _
        "                             - Edsger Dijkstra 1930..2002" & _
        vbNewLine & vbNewLine & vbNewLine & _
        """But the man who knows the relation between the forces of nature " & _
        "and action, sees how some forces of Nature work upon other forces " & _
        "of Nature, and becomes not their slave.""" & _
        vbNewLine & vbNewLine & _
        "                             - Bhagavad-Gita" & _
        vbNewLine & vbNewLine & vbNewLine & _
        """I could be bound in a nutshell and count myself king of absolute " & _
        "space.""" & _
        vbNewLine & vbNewLine & _
        "                             - Shakespeare, Hamlet"
    ' --- Class name
    Private Const CLASS_NAME As String = "quickbasicEngine"
    ' --- Inspection
    Private Const INSPECTION_USABLE As String = _
            "The quickBasicEngine object must be usable"
    Private Const INSPECTION_SCANNEROK As String = _
            "The scanner object must pass its own inspection"
    Private Const INSPECTION_MAXCOLLECTIONITEMS As Byte = 100
    Private Const INSPECTION_POLISHOK As String = _
            "The Polish collection must be a collection of qbPolish objects, " & _
            "which pass their own inspection. If there are many such objects, " & _
            "a random inspection is made of a maximum number of objects."
    Private Const INSPECTION_VARIABLESOK As String = _
            "The Variables collection must be a collection of qbVariable objects, " & _
            "which pass their own inspection. If there are many such objects, " & _
            "a random inspection is made of a maximum number of objects." & _
            vbNewLine & vbNewLine & _
            "Each qbVariable's VariableName must point as a collection key to the " & _
            "same location as the qbVariable: each qbVariable's Tag must contain " & _
            "the variable index."
    Private Const INSPECTION_SUBFUNCTIONINDEX As String = _
            "Each entry in the subroutine/function index collection must " & _
            "be a Collection containing two items: item(1) must be a subroutine " & _
            "or function name and item(2) must be its index in the main table " & _
            "of subroutines and functions."
    Private Const INSPECTION_CONSTANTEXPRESSIONINDEX As String = _
            "Each entry in the constant expression index collection must " & _
            "be a Collection containing two items: item(1) must be a subroutine " & _
            "or function name and item(2) must be its index in the main table " & _
            "of constant expressions."
    ' --- Postfix data types (Hergert 1989) 
    Private Const POSTFIX_DATATYPE_CHARS As String = "%&!#$"
    Private Const POSTFIX_DATATYPE_CHAR_SHORT As String = "%"
    Private Const POSTFIX_DATATYPE_CHAR_LONG As String = "&"
    Private Const POSTFIX_DATATYPE_CHAR_SINGLE As String = "!"
    Private Const POSTFIX_DATATYPE_CHAR_DOUBLE As String = "#"
    Private Const POSTFIX_DATATYPE_CHAR_STRING As String = "$"
    ' --- Tests
    Private Const TEST_CODE_NFACTORIAL As String = _
    "' ***** CALCULATION OF N FACTORIAL *****&#00013&#00010" & _
    "Dim N&#00013&#00010" & _
    "Dim F&#00013&#00010" & _
    "PRINT ""ENTER N""&#00013&#00010" & _
    "INPUT N&#00013&#00010" & _
    "IF N<>INT(N) THEN&#00013&#00010" & _
    "PRINT ""N VALUE "" & N & "" IS NOT AN INTEGER""&#00013&#00010" & _
    "END&#00013&#00010" & _
    "END IF&#00013&#00010" & _
    "IF N<=0 THEN&#00013&#00010" & _
    "PRINT ""N VALUE "" & N & "" IS NOT A POSITIVE NUMBER""&#00013&#00010" & _
    "END&#00013&#00010" & _
    "END IF&#00013&#00010" & _
    "F = 1&#00013&#00010" & _
    "Dim N2&#00013&#00010" & _
    "FOR N2 = N TO 2 STEP -1&#00013&#00010" & _
    "F = F * N2&#00013&#00010" & _
    "NEXT N2&#00013&#00010" & _
    "PRINT ""THE FACTORIAL OF "" & N & "" IS "" & F "
    Private Const TEST_CODE_DIMSTMT As String = _
    "Dim booFlag as Boolean&#00013&#00010" & _
    "booFlag = True&#00013&#00010" & _
    "Dim bytSmallInteger As Byte&#00013&#00010" & _
    "bytSmallInteger = 255&#00013&#00010" & _
    "Dim intMediumInteger As Integer&#00013&#00010" & _
    "intMediumInteger = -32768&#00013&#00010" & _
    "Dim lngLargeInteger As Long&#00013&#00010" & _
    "lngLargeInteger = 32768&#00013&#00010" & _
    "Dim sngSingleFP As Single&#00013&#00010" & _
    "sngSingleFP = 1.1&#00013&#00010" & _
    "Dim dblDoubleFP As Double&#00013&#00010" & _
    "dblDoubleFP = 1.00000001&#00013&#00010" & _
    "Dim strText As String&#00013&#00010" & _
    "strText = ""Text""&#00013&#00010" & _
    "Print booFlag&#00013&#00010" & _
    "Print bytSmallInteger&#00013&#00010" & _
    "Print intMediumInteger&#00013&#00010" & _
    "Print lngLargeInteger&#00013&#00010" & _
    "Print sngSingleFP&#00013&#00010" & _
    "Print dblDoubleFP&#00013&#00010" & _
    "Print strText"
    Private Const TEST_CODE As String = _
    "Expression 1+1" & vbNewLine & "2" & _
    vbNewLine & vbNewLine & _
    "Expression ""Ooga"" & ""Chukka""" & vbNewLine & "OogaChukka" & _
    vbNewLine & vbNewLine & _
    "Expression 5478/3+21-((4+1)*8) + .1" & vbNewLine & "1807.10000000149" & _
    vbNewLine & vbNewLine & _
    "Program Print(""Hello world"")" & vbNewLine & " " & vbNewLine & "Hello world&#00013&#00010" & _
    vbNewLine & vbNewLine & _
    "Program " & TEST_CODE_NFACTORIAL & vbNewLine & "5" & vbNewLine & _
    "ENTER N&#00013&#00010THE FACTORIAL OF 5 IS 120&#00013&#00010" & _
    vbNewLine & vbNewLine & _
    "Program a = ""Hello world"": Print a" & vbNewLine & " " & vbNewLine & "Hello world&#00013&#00010" & _
    vbNewLine & vbNewLine & _
    "Expression 1& & ""1""" & vbNewLine & "11" & _
    vbNewLine & vbNewLine & _
    "Expression 1& + ""1""" & vbNewLine & "2" & _
    vbNewLine & vbNewLine & _
    "Expression (((((400-32-40+89+901)/4/2*8)+3+3+3)/40-21+9)-5-6-0-0-0-0-0)*1*1*1" & vbNewLine & "10.1749992370605" & _
    vbNewLine & vbNewLine & _
    "Program " & TEST_CODE_DIMSTMT & vbNewLine & " " & vbNewLine & _
    "True&#00013&#00010255&#00013&#00010-32768&#00013&#0001032768&#00013&#000101.1&#00013&#000101.00000001&#00013&#00010Text&#00013&#00010"
    ' --- Reserved words
    Private Const QB_RESERVED_WORDS As String = _
    "abs and andalso as asc boolean byref byval byte ceil chr circle cos data dim do " & _
    "double else end endif eval exit explicit extension false floor for function gosub goto if iif " & _
    "input int integer isnumeric lbound lcase left len let like log long loop max " & _
    "mid min next not option or orelse print randomize read rem replace return right rnd " & _
    "screen sgn sin single step stop strng sub then to trace time true " & _
    "until ubound ucase until utility variant while wend"
    ' --- Miscellaneous constants
    Private Const MAX_UTILITY_ARGS As Integer = 10
    Private Const CONSTANTEXPRESSION_MAXCACHE As Integer = 1024
    Private Const INTERPRETER_STACK_BLOCK As Integer = 32
    Private Const INTERPRETER_STACK_MAXFREE As Integer = INTERPRETER_STACK_BLOCK * 10

    ' ***** Events ****************************************************
    ' *                                                               *
    ' * When adding an event note that you should add a Case for the  *
    ' * event to the raiseEvent_ method.                              *
    ' *                                                               *
    ' *****************************************************************
    Public Event codeChangeEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                 ByVal objOp As qbPolish.qbPolish, _
                                 ByVal intOpIndex As Integer)
    Public Event codeGenEvent(ByVal objQBsender As qbQuickBasicEngine, _
                              ByVal objOp As qbPolish.qbPolish)
    Public Event codeRemoveEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                 ByVal intOpIndex As Integer)
    Public Event compileErrorEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                   ByVal strMessage As String, _
                                   ByVal intIndex As Integer, _
                                   ByVal intContextLength As Integer, _
                                   ByVal intLinenumber As Integer, _
                                   ByVal strHelp As String, _
                                   ByVal strCode As String)
    Public Event interpretClsEvent(ByVal objQBsender As qbQuickBasicEngine)
    Public Event interpretErrorEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                     ByVal strMessage As String, _
                                     ByVal intIndex As Integer, _
                                     ByVal strHelp As String)
    Public Event interpretInputEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                     ByRef strChars As String)
    Public Event interpretPrintEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                     ByVal strOutstring As String)
#If QUICKBASICENGINE_POPCHECK Then
    Public Event interpretTraceEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                     ByVal intIndex As Integer, _
                                     ByVal objStack As Stack, _
                                     ByVal colStorage As Collection)
#Else
    Public Event interpretTraceFastEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                         ByVal intIndex As Integer, _
                                         ByVal objStack() As qbVariable.qbVariable, _
                                         ByVal intStackTop As Integer, _
                                         ByVal colStorage As Collection)
#End If
    Public Event loopEvent(ByVal objQBsender As qbQuickBasicEngine, _
                           ByVal strActivity As String, _
                           ByVal strEntity As String, _
                           ByVal intNumber As Integer, _
                           ByVal intCount As Integer, _
                           ByVal intLevel As Integer, _
                           ByVal strComment As String)
    Public Event msgEvent(ByVal objQBsender As qbQuickBasicEngine, _
                          ByVal strMessage As String)
    Public Event parseEvent(ByVal objQBsender As qbQuickBasicEngine, _
                            ByVal strGrammarCategory As String, _
                            ByVal booTerminal As Boolean, _
                            ByVal intSrcStartIndex As Integer, _
                            ByVal intSrcLength As Integer, _
                            ByVal intTokStartIndex As Integer, _
                            ByVal intTokLength As Integer, _
                            ByVal intObjStartIndex As Integer, _
                            ByVal intObjLength As Integer, _
                            ByVal strComment As String, _
                            ByVal intLevel As Integer)
    Public Event parseFailEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                ByVal strGrammarCategory As String)
    Public Event parseStartEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                 ByVal strGrammarCategory As String)
    Public Event scanEvent(ByVal objQBsender As qbQuickBasicEngine, _
                           ByVal objToken As qbToken.qbToken)
    Public Event threadStatusChangeEvent(ByVal objQBsender As qbQuickBasicEngine)
    Public Event userErrorEvent(ByVal objQBsender As qbQuickBasicEngine, _
                                ByVal strDescription As String, _
                                ByVal strHelp As String)

    ' ***** Compilation status *****
    Private Enum ENUstatus
        initial
        scanned
        compiled
        assembled
    End Enum

    ' ***** Source code type *****    
    Private Enum ENUsourceCodeType
        unknown                     ' Determine by trying to compile as program, 
                                    ' then as an immediate command
        immediateCommand            ' Expression
        program                     ' Complete program
        invalid                     ' Error value
    End Enum
    Private Const SOURCECODETYPELIST As String = "unknown immediateCommand program"

    ' ***** Withevents object *****
    Private WithEvents OBJqbeWithEvents As quickBasicEngine.qbQuickBasicEngine

    ' ***** lValue parsing return codes *****
    ' Have found variable whose address is on the stack
    Private Const MEMORY_LOCATION_STACKEDADDRESS As Integer = -1
    ' Array reference is on the stack
    Private Const MEMORY_LOCATION_STACKFRAME As Integer = -2

    ' ***** Object Constructor ****************************************

    Public Sub New()
        Try
            OBJstate = New stateClass
        Catch OBJexception As Exception
            errorHandler_("Cannot create stateClass: " & Err.Number & " " & Err.Description, _
                          "New", _
                          "Object won't be usable", _
                          objException)
            Return
        End Try
        With OBJstate.usrState
            Try
                OBJthreadStatus = New threadStatus
            Catch objException As Exception
                errorHandler_("Cannot create threadStatus: " & Err.Number & " " & Err.Description, _
                              "New", _
                              "Object won't be usable", _
                              objException)
                Return
            End Try
            SyncLock OBJstate
                .strName = ClassName & _
                           _OBJutilities.alignRight(CStr(Interlocked.Increment(_INTsequence)), 4, "0") & " " & _
                           Now
                Try
                    OBJcollectionUtilities = New collectionUtilities.collectionUtilities
                    .objCollectionUtilities = OBJcollectionUtilities
                Catch objException As Exception
                    errorHandler_("Unable to create reference objects in state: " & _
                                Err.Number & " " & Err.Description, _
                                "New", _
                                "Object will not be usable", _
                                objException)
                End Try
                .booAssemblyRemovesCode = True
#If QUICKBASICENGINE_EXTENSION Then
                .booGenerateNOPs = True
#End If
                OBJthreadStatus.mkReady()
                .booUsable = True
                .booUsable = inspection_()
            End SyncLock
        End With
    End Sub

    ' ***** Public procedures *****************************************
    ' *                                                               *
    ' * Also may contain Private procedures that do tasks on behalf of*
    ' * a single Public procedure.                                    *
    ' *                                                               *
    ' *****************************************************************

    ' -----------------------------------------------------------------
    ' Return Easter Egg
    '
    '
    Public Shared ReadOnly Property About() As String
        Get
            Return ABOUT_INFO
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Assemble compiled code
    '
    '
    Public Function assemble() As Boolean
        Return (CBool(dispatch_("assemble", True, False, "Returning False")))
    End Function

    ' -----------------------------------------------------------------
    ' Tell caller if code has been assembled
    '
    '
    ' Note that this method may NOT be called internally
    '
    '
    Public Function assembled() As Boolean
        Return (CBool(dispatch_("assembled", True, False, "Returning False")))
    End Function

    ' ----------------------------------------------------------------------
    ' Return True (always assemble) or False
    '
    '
    Public Property AssemblyRemovesCode() As Boolean
        Get
            Return (CBool(dispatch_("AssemblyRemovesCode Get", _
                                    True, _
                                    False, _
                                    "Returning False")))
        End Get
        Set(ByVal booNewValue As Boolean)
            dispatch_("AssemblyRemovesCode Set", _
                      True, _
                      Nothing, _
                      "No change made", _
                      booNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return the name of this class
    '
    '
    Public Shared ReadOnly Property ClassName() As String
        Get
            Return CLASS_NAME
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Clear the compiler object to empty values
    '
    '
    Public Function clear() As Boolean
        Return (CBool(dispatch_("clear", _
                                True, _
                                False, _
                                "Returning False")))
    End Function

    ' -----------------------------------------------------------------
    ' Clear storage to default values
    '
    '
    Public Function clearStorage() As Boolean
        Return (CBool(dispatch_("clearStorage", _
                                True, _
                                False, _
                                "Returning False")))
    End Function

    ' -----------------------------------------------------------------
    ' Clone the quickBasicEngine
    '
    '
    Public Function clone() As quickBasicEngine.qbQuickBasicEngine
        Return (CType(dispatch_("clone", _
                                True, _
                                False, _
                                "Returning False"), _
                      quickBasicEngine.qbQuickBasicEngine))
    End Function

    ' -----------------------------------------------------------------
    ' Return the code type
    '
    '
    Public Shared Function codeType(ByVal strCode As String) As String
        Try
            Dim objValue As Object = eval(strCode)
            Return ENUsourceCodeType.immediateCommand.ToString
        Catch
            Try
                Dim objRunner As quickBasicEngine.qbQuickBasicEngine
                Try
                    objRunner = New quickBasicEngine.qbQuickBasicEngine
                Catch
                    errorHandler_("Cannot create Quick Basic engine: " & _
                                  Err.Number & " " & Err.Description, _
                                  "codeType", _
                                  "Returning ""invalid""")
                    Return ENUsourceCodeType.invalid.ToString
                End Try
                If Not objRunner.run(strCode) Then Return ("invalid")
                Return ENUsourceCodeType.program.ToString
            Catch
                Return ENUsourceCodeType.invalid.ToString
            End Try
        End Try
    End Function    

    ' -----------------------------------------------------------------
    ' Compare the quickBasic engine
    '
    '
    ' --- Public interface
    Public Function compareTo(ByVal objQBE As quickBasicEngine.qbQuickBasicEngine) As Boolean
        Return (CBool(dispatch_("compareTo", _
                                True, _
                                False, _
                                "Returning False", _
                                objQBE)))
    End Function

    ' -----------------------------------------------------------------
    ' Compile the source code
    '
    '
    Public Function compile() As Boolean
        Return (CBool(dispatch_("compile", _
                                True, _
                                False, _
                                "Returning False")))
    End Function

    ' -----------------------------------------------------------------
    ' Tell caller if source code has been compiled
    '
    '
    Public Function compiled() As Boolean
        Return (CBool(dispatch_("compiled", _
                                True, _
                                False, _
                                "Returning False")))
    End Function

    ' -----------------------------------------------------------------
    ' Return and assign the constant folding property
    '
    '
    Public Property ConstantFolding() As Boolean
        Get
            Return (CBool(dispatch_("constantFolding Get", _
                                    True, _
                                    False, _
                                    "Returning False")))
        End Get
        Set(ByVal booNewValue As Boolean)
            dispatch_("constantFolding Set", _
                      True, _
                      Nothing, _
                      "No change made", _
                      booNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return and assign the "degenerate" op removal property
    '
    '
    Public Property DegenerateOpRemoval() As Boolean
        Get
            Return CBool(dispatch_("DegenerateOpRemoval Get", True, False, "Returning False"))
        End Get
        Set(ByVal booNewValue As Boolean)
            dispatch_("DegenerateOpRemoval Set", True, Nothing, "No change made", booNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Dispose of reference objects and mark unusable
    '
    '
    Public Function dispose() As Boolean
        Return (dispose(True))
    End Function
    Public Function dispose(ByVal booInspect As Boolean) As Boolean
        Return CBool(dispatch_("dispose", True, False, "Returning False", booInspect))
    End Function

    ' ----------------------------------------------------------------------
    ' Return dedicatory Easter Egg
    '
    '
    Public Shared ReadOnly Property EasterEgg() As String
        Get
            Return About & _
                   vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
                   EASTEREGG_INFO
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Evaluate an immediate command (lightweight method)
    '
    '
    ' --- No event log  
    Public Overloads Shared Function eval(ByVal strExpression As String) _
           As qbVariable.qbVariable
        Dim colDummy As Collection
        Return (eval_(strExpression, False, colDummy))
    End Function
    ' --- Event log returned
    Public Overloads Shared Function eval(ByVal strExpression As String, _
                                          ByRef strEventLog As String) _
           As qbVariable.qbVariable
        Dim colEventLog As Collection
        Return (eval_(strExpression, True, colEventLog))
    End Function
    ' --- Return string or error report and error indicator
    Public Overloads Shared Function eval(ByVal strExpression As String, _
                                          ByRef booError As Boolean) _
           As String
        Dim objValue As qbVariable.qbVariable
        Dim colEventLog As Collection
        objValue = eval_(strExpression, True, colEventLog)
        Dim strErrors As String = eventLog2ErrorList(colEventLog)
        If strErrors <> "" Then
            booError = True
            Return (strErrors)
        End If
        Dim strValue As String
        Try
            strValue = CStr(objValue.value)
        Catch
            booError = True
            Return ("Cannot convert eval's value to a string")
        End Try
        booError = False
        Return (strValue)
    End Function
    ' --- Common logic  
    Private Shared Function eval_(ByVal strExpression As String, _
                                  ByVal booEventLog As Boolean, _
                                  ByRef colEventLog As Collection) _
           As qbVariable.qbVariable
        Dim objQuickBasicEngine As quickBasicEngine.qbQuickBasicEngine
        Try
            objQuickBasicEngine = New quickBasicEngine.qbQuickBasicEngine
            Dim objValue As qbVariable.qbVariable
            With objQuickBasicEngine
                .EventLogging = booEventLog
                objValue = .evaluate(strExpression)
                If booEventLog Then
                    colEventLog = .EventLog
                End If
                .dispose()
            End With
            Return (objValue)
        Catch
            errorHandler_("Error occured in evaluating the expression " & _
                          _OBJutilities.enquote(strExpression) & ": " & _
                          Err.Number & " " & Err.Description, _
                          "eval", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Evaluate an immediate command
    '
    '
    ' --- No event log
    Public Overloads Function evaluate(ByVal strExpression As String) As qbVariable.qbVariable
        Dim strDummy As String
        Return (evaluate__(strExpression, False, strDummy))
    End Function
    ' --- Place result in a .Net value
    Public Overloads Sub evaluate(ByVal strExpression As String, _
                                  ByRef objValue As Object)
        Dim strDummy As String
        objValue = evaluate__(strExpression, False, strDummy).value
    End Sub
    ' --- Return event log
    Public Overloads Function evaluate(ByVal strExpression As String, _
                                       ByRef strEventLog As String) As qbVariable.qbVariable
        Return (evaluate__(strExpression, True, strEventLog))
    End Function
    ' --- Common logic
    Private Overloads Function evaluate__(ByVal strExpression As String, _
                                          ByVal booEventLog As Boolean, _
                                          ByRef strEventLog As String) As qbVariable.qbVariable
        Return CType(dispatch_("Evaluate", _
                                strEventLog, _
                                Nothing, _
                                "Returning Nothing", _
                                strExpression, _
                                booEventLog), _
                     qbVariable.qbVariable)
    End Function

    ' -----------------------------------------------------------------------
    ' Return the value of the most recent evaluated expression
    '
    '
    Public ReadOnly Property Evaluation() As qbVariable.qbVariable
        Get
            Return CType(dispatch_("Evaluation get", _
                                   True, _
                                   Nothing, _
                                   "Returning Nothing"), _
                         qbVariable.qbVariable)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Return the value of the most recent evaluated expression
    '
    '
    Public Function evaluationValue() As Object
        Return dispatch_("evaluationValue", True, Nothing, "Returning Nothing")
    End Function

    ' -----------------------------------------------------------------------
    ' Return the event log
    '
    '
    Public ReadOnly Property EventLog() As Collection
        Get
            Return (CType(dispatch_("EventLog Get", _
                                    True, _
                                    Nothing, _
                                    "Returning Nothing"), _
                          Collection))
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Convert an event log to a compiler and interpreter error list
    '
    '
    ' --- Stateful overload
    Public Overloads Function eventLog2ErrorList() As String
        Return CStr(dispatch_("eventLog2ErrorList", _
                                True, _
                                "", _
                                "Return a null string"))
    End Function
    ' --- Stateless overload
    Public Overloads Shared Function eventLog2ErrorList(ByVal colEventLog As Collection) _
           As String
        Return (eventLog2ErrorList_(colEventLog))
    End Function

    ' -----------------------------------------------------------------
    ' Format the event log
    '
    '
    ' Note that this method uses RaiseEvent directly rather than calling
    ' the raiseEvent_ method to avoid recursive expansion of the event
    ' collection as it is reporting events.
    '
    '
    ' --- Format the complete log
    Public Overloads Function eventLogFormat() As String
        Return (eventLogFormat(1, OBJstate.usrState.colEventLog.Count))
    End Function
    ' --- Specifies start index
    Public Overloads Function eventLogFormat(ByVal intStartIndex As Integer) _
           As String
        Return (eventLogFormat(intStartIndex, _
                                  OBJstate.usrState.colEventLog.Count _
                                  - _
                                  intStartIndex _
                                  + _
                                  1))
    End Function
    ' --- Shared version: formats a complete log passed as a parameter
    Public Overloads Shared Function eventLogFormat(ByVal colEventLog As Collection) _
           As String
        Return (eventLogFormat_(colEventLog, 1, colEventLog.Count))
    End Function
    ' --- Specifies start index and count
    Public Overloads Function eventLogFormat(ByVal intStartIndex As Integer, _
                                             ByVal intCount As Integer) _
           As String
        Return CStr(dispatch_("eventLogFormat", _
                                True, _
                                "", _
                                "Return a null string", _
                                intStartIndex, _
                                intCount))
    End Function

    ' -----------------------------------------------------------------------
    ' Return and change event logging status
    '
    '
    Public Property EventLogging() As Boolean
        Get
            Return CBool(dispatch_("EventLogging Get", True, False, "Return false"))
        End Get
        Set(ByVal booNewValue As Boolean)
           dispatch_("EventLogging Set", True, Nothing, "No change made", booNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Return setting of QUICKBASICENGINE_EXTENSION property at run time
    '
    ' 
    Public Shared ReadOnly Property ExtensionAvailable() As Boolean
        Get
#If QUICKBASICENGINE_EXTENSION Then
            Return True
#End If
        End Get
    End Property

#If QUICKBASICENGINE_EXTENSION Then

    ' -----------------------------------------------------------------
    ' Return and assign generate NOPs/REMs property
    '
    '
    Public Property GenerateNOPs() As Boolean
    Get
        Return CBool(dispatch_("GenerateNOPs Get", True, False, "Return false"))
    End Get
    Set(ByVal booNewValue As Boolean)
        dispatch_("GenerateNOPs Set", True, Nothing, "No change made", booNewValue)
    End Set
    End Property

#End If

    ' -----------------------------------------------------------------
    ' Get the run status
    '
    '
    Public Overloads Function getThreadStatus() As String
        Return (OBJthreadStatus.getThreadStatus(True))
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the object
    '
    '
    Public Overloads Function inspect() As Boolean
        Return CBool(dispatch_("inspect", True, False, "Return False"))
    End Function
    Public Overloads Function inspect(ByRef strReport As String) As Boolean
        Return CBool(dispatch_("inspect", strReport, False, "Return False"))
    End Function

    ' ----------------------------------------------------------------------
    ' Set and return option to inspect the compiler and its objects
    '
    '
    Public Property InspectCompilerObjects() As Boolean
    Get
        Return CBool(dispatch_("InspectCompilerObjects Get", True, False, "Returning False"))
    End Get
    Set(ByVal booNewValue As Boolean)
        dispatch_("InspectCompilerObjects Set", True, Nothing, "No change made", booNewValue)
    End Set
    End Property

    ' ----------------------------------------------------------------------
    ' Interpret the code
    '
    '
    Public Function interpret() As Object
        Return dispatch_("interpret", True, Nothing, "Returning Nothing")
    End Function

    ' -----------------------------------------------------------------
    ' Make the instance not usable
    '
    '
    Public Function mkUnusable() As Boolean
        Return CBool(dispatch_("mkUnusable", True, False, "Returning False"))
    End Function

    ' -----------------------------------------------------------------
    ' Translate if possible to IL and run
    '
    '
    Public Function msilRun() As Object
        Return dispatch_("msilRun", True, Nothing, "Returning Nothing")
    End Function

    ' -----------------------------------------------------------------
    ' Return and assign object's name
    '
    '
    Public Property Name() As String
        Get
            dispatch_("Name Get", True, "", "Returning Null string")
        End Get
        Set(ByVal strNewValue As String)
            dispatch_("Name Set", True, Nothing, "No change made")
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Convert the state of the object to eXtended Markup Language
    '
    '
    Public Function object2XML(Optional ByVal booAboutComment As Boolean = False, _
                               Optional ByVal booStateComment As Boolean = True) _
           As String
        Return CStr(dispatch_("object2XML", True, "", "Returning null string", _
                                booAboutComment, _
                                booStateComment))
    End Function                                          

    ' ----------------------------------------------------------------------
    ' Return the Polish collection
    '
    '
    Public ReadOnly Property PolishCollection() As Collection
        Get
            Return CType(dispatch_("PolishCollection Get", True, Nothing, "Returning nothing"), _
                         Collection)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return and assign PopCheck option
    '
    '
    Public Shared ReadOnly Property PopCheckAvailable() As Boolean
        Get
#If QUICKBASICENGINE_POPCHECK Then
            Return true
#Else
            Return False
#End If
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Resets the quickBasicEngine
    '
    '
    Public Function reset() As Boolean
        Return CBool(dispatch_("reset", True, False, "Returning False"))
    End Function

    ' ----------------------------------------------------------------------
    ' Resume the engine
    '
    '
    Public Function resumeQBE() As Boolean
        Return OBJthreadStatus.setThreadStatus _
               (threadStatus.ENUthreadStatus.Ready)
    End Function

    ' ----------------------------------------------------------------------
    ' Run source code
    '
    '
    Public Overloads Function run() As Boolean
        Return (run("program"))
    End Function
    Public Overloads Function run(ByVal strSourceCodeType As String) As Boolean
        Return CBool(dispatch_("run", True, False, "Returning False", strSourceCodeType))
    End Function

    ' -----------------------------------------------------------------
    ' Get the thread count
    '
    '
    Public Function runningThreads() As Integer
        Return (OBJthreadStatus.runningThreads)
    End Function

    ' -----------------------------------------------------------------------
    ' Scan the code
    '
    '
    Public Function scan() As Boolean
        Return CBool(dispatch_("scan", True, False, "Returning False"))
    End Function

    ' -----------------------------------------------------------------------
    ' Return True if code is scanned
    '
    '
    Public Function scanned() As Boolean
        Return CBool(dispatch_("scanned", True, False, "Returning False"))
    End Function

    ' -----------------------------------------------------------------------
    ' Return the scanner object
    '
    '
    Public ReadOnly Property Scanner() As qbScanner.qbScanner
        Get
            Return CType(dispatch_("Scanner Get", True, Nothing, "Returning Nothing"), _
                         qbScanner.qbScanner)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Assign and return source code
    '
    '
    ' Because the source code is stored in our scanner object, this method
    ' may have to create the scanner whenever the source code is assigned
    ' prior to the first scan.
    '
    '
    Public Property SourceCode() As String
        Get
            Return CStr(dispatch_("SourceCode Get", True, "", "Returning null string"))
        End Get
        Set(ByVal strNewValue As String)
            dispatch_("SourceCode Set", True, Nothing, "No change made", strNewValue)
        End Set
    End Property

#If QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Serializes the interpreter's stack when it is a PopChecked Stack
    '
    '
    Public Overloads Shared Function stack2String(ByRef objStack As Stack) _
           As String
        With objStack
            Dim colSave As Collection
            Try
                colSave = New Collection
            Catch objException As Exception
                errorHandler__("Cannot create collection: " & _
                               ClassName, _
                               Err.Number & " " & Err.Description, _
                               "stack2String", _
                               "Returning null string", _
                               objException)
                Return ("")
            End Try
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                Try
                    colSave.Add(.Pop)
                Catch objException As Exception
                    errorHandler__("Cannot save stack: " & _
                                   ClassName, _
                                   Err.Number & " " & Err.Description, _
                                   "stack2String", _
                                   "Returning null string", _
                                   objException)
                    Return ("")
                End Try
            Next intIndex1
            Dim objNext As qbVariable.qbVariable
            Dim strNext As String
            Dim strOutstring As String
            With colSave
                For intIndex1 = 1 To .Count
                    objNext = CType(.Item(intIndex1), qbVariable.qbVariable)
                    _OBJutilities.append(strOutstring, _
                                         vbNewLine, _
                                         stackEntry2String_(CType(colSave.Item(intIndex1), _
                                                                  qbVariable.qbVariable)))
                Next intIndex1
            End With
            For intIndex1 = colSave.Count To 1 Step -1
                Try
                    .Push(colSave.Item(intIndex1))
                Catch objException As Exception
                    errorHandler__("Cannot restore stack: " & _
                                   ClassName, _
                                   Err.Number & " " & Err.Description, _
                                   "stack2String", _
                                   "Returning null string", _
                                   objException)
                    Return ("")
                End Try
            Next intIndex1
            Return (strOutstring)
        End With
    End Function
#End If

#If Not QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Serializes the interpreter's stack when it is not a PopChecked Stack
    '
    '
    Public Overloads Shared Function stack2StringFast(ByRef objStack() As qbVariable.qbVariable, _
                                                      ByVal intStackTop As Integer) _
           As String
        Dim intIndex1 As Integer
        Dim strNext As String
        Dim strOutstring As String
        For intIndex1 = intStackTop To 1 Step -1
            _OBJutilities.append(strOutstring, _
                                 vbNewLine, _
                                 stackEntry2String_(objStack(intIndex1)))
        Next intIndex1
        Return (strOutstring)
    End Function
#End If

    ' ----------------------------------------------------------------------
    ' Convert stack entry to serialized form
    '
    '
    Private Shared Function stackEntry2String_(ByVal objEntry As qbVariable.qbVariable) _
            As String
        Dim strOutstring As String
        With objEntry
            If .Dope.isScalar Then
                strOutstring = .ToString
            Else
                strOutstring = .Dope.ToString
            End If
            strOutstring = _OBJutilities.ellipsis(strOutstring, 32)
            If (TypeOf .Tag Is Integer) Then
                strOutstring = "*->[" & strOutstring & "]@" & CType(.Tag, Integer)
            End If
            Return strOutstring
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Stop the engine
    '
    '
    Public Function stopQBE() As Boolean
        SyncLock OBJthreadStatus
            OBJthreadStatus.stopThreads()
        End SyncLock
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Serializes the interpreter's storage
    '
    '
    ' --- Basic call
    Public Overloads Shared Function storage2String(ByVal colStorage As Collection) _
           As String
        Return storage2String(colStorage, False)
    End Function
    ' --- Exposes option to include array values
    Public Overloads Shared Function storage2String(ByVal colStorage As Collection, _
                                                    ByVal booIncludeArrayValues As Boolean) _
           As String
        With colStorage
            Dim intIndex1 As Integer
            Dim objNext As qbVariable.qbVariable
            Dim strNext As String
            Dim strOutString As String
            For intIndex1 = 1 To .Count
                objNext = CType(.Item(intIndex1), qbVariable.qbVariable)
                With objNext
                    If .Dope.isArray AndAlso Not booIncludeArrayValues Then
                        strNext = .Dope.ToString
                    Else
                        strNext = .ToString
                    End If
                End With
                _OBJutilities.append(strOutString, _
                                     vbNewLine, _
                                     _OBJutilities.alignRight(CStr(intIndex1), _
                                                              CInt(Math.Log10(.Count) + 1), _
                                                              "0") & " " & _
                                     objNext.VariableName & " " & strNext)
            Next intIndex1
            Return (strOutString)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return and assign the Tag
    '
    '
    Public Property Tag() As Object
        Get
            Return (dispatch_("Tag Get", True, Nothing, "Returning Nothing"))
        End Get
        Set(ByVal objNewValue As Object)
            dispatch_("Tag Set", True, Nothing, "No change made", objNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Tester
    '
    '
    Public Overloads Function test(ByRef strReport As String) As Boolean
        Return test(strReport, False)
    End Function    
    Public Overloads Function test(ByRef strReport As String, _
                                   ByVal booEventLog As Boolean) As Boolean
        dispatch_("test", strReport, False, "Returning False", booEventLog)
    End Function

    ' -----------------------------------------------------------------
    ' Return or set the indexed or named variable
    '
    '
    Public Property Variable(ByVal objIndex As Object) _
           As qbVariable.qbVariable
        Get
            Variable = CType(dispatch_("Variable Get", _
                                       True, _
                                       Nothing, _
                                       "Returning Nothing", _
                                       objIndex), _
                                       qbVariable.qbVariable)
        End Get
        Set(ByVal objNewValue As qbVariable.qbVariable)
             dispatch_("Variable Set", _
                       True, _
                       Nothing, _
                       "Returning Nothing", _
                       objIndex, _
                       objNewValue)
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return the indexed or named variable
    '
    '
    Public ReadOnly Property VariableCount() As Integer
        Get
            VariableCount = CInt(dispatch_("VariableCount Get", _
                                            True, _
                                            0, _
                                            "Returning zero"))
        End Get
    End Property

    ' ***** Friend procedures *****************************************

    ' -----------------------------------------------------------------
    ' Return and assign the qbScanner
    '
    '
    Friend Property qbScanner() As qbScanner.qbScanner
        Get
            Return CType(dispatch_("qbScanner Get", True, Nothing, "Returning Nothing"), _
                         qbScanner.qbScanner)
        End Get
        Set(ByVal objNewValue As qbScanner.qbScanner)
            dispatch_("qbScanner Set", True, Nothing, "No change made", objNewValue)
        End Set
    End Property

    ' ***** Private procedures ****************************************

    ' ----------------------------------------------------------------------
    ' Assemble compiled code
    '
    '
    Private Function assemble_() As Boolean
        With OBJstate.usrState
            If Not checkUsable_("assemble", "Returning False") Then Return (False)
            If Not .objScanner.Scanned Then
                If Not .objScanner.scan Then Return (False)
            End If
            If Not .booCompiled Then
                If Not compile_() Then Return (False)
            End If
            If assembler_() Then
                .booAssembled = True
                Return (True)
            End If
        End With
    End Function

    ' ----------------------------------------------------------------- 
    ' Assemble code
    '
    '
    Private Function assembler_() As Boolean
        Dim colEntry As Collection
        Dim colNewLabels As Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim intIndex4 As Integer
        Dim intDelete As Integer
        Dim intDeleted As Integer
        Dim objPolishHandle As qbPolish.qbPolish
        Dim strKey As String
        With OBJstate.usrState
            ' --- Create the label table
            Try
                .colLabel = New Collection
                colNewLabels = New Collection
            Catch objException As Exception
                errorHandler_("Cannot create label tables: " & _
                              Err.Number & " " & Err.Description, _
                              "assembler_", _
                              "Returning False", _
                              objException)
            End Try
            ' --- Pass one: get label definitions while removing non-executable statements  
            intIndex1 = 1
            Do While intIndex1 <= .colPolish.Count
                For intIndex2 = intIndex1 To .colPolish.Count
                    If loopEventInterface_("Assembling code in pass one", _
                                           "Opcode", _
                                           intIndex2, _
                                           .colPolish.Count, _
                                           0, _
                                           .colLabel.Count & " labels have been found") Then
                        Exit For
                    End If
                    With CType(.colPolish.Item(intIndex2), qbPolish.qbPolish)
                        If .Opcode = ENUop.opLabel Then
                            strKey = object2Key_(.Operand)
                            Try
                                If OBJstate.usrState.booAssemblyRemovesCode Then
                                    ' Label's position is unknown
                                    colNewLabels.Add(strKey)
                                Else
                                    ' Label's position is known
                                    colEntry = New Collection
                                    With colEntry
                                        .Add(strKey) : .Add(intIndex2)
                                    End With
                                    colNewLabels.Add(colEntry)
                                End If
                            Catch objException As Exception
                                errorHandler_("Internal compiler error occured in new label table extension: " & _
                                              Err.Number & " " & Err.Description, _
                                              "assembler_", _
                                              "Marking object unusable: returning False", _
                                              objException)
                                OBJstate.usrState.booUsable = False
                                Return (False)
                            End Try
                        End If
                        If OBJstate.usrState.booAssemblyRemovesCode _
                           AndAlso _
                           .Opcode <> ENUop.opLabel _
                           AndAlso _
                           .Opcode <> ENUop.opNop _
                           AndAlso _
                           .Opcode <> ENUop.opRem Then
                            Exit For
                        End If
                    End With
                Next intIndex2
                intIndex3 = 0
                If OBJstate.usrState.booAssemblyRemovesCode Then
                    For intIndex1 = intIndex2 - 1 To intIndex1 Step -1
                        objPolishHandle = CType(.colPolish.Item(intIndex1), qbPolish.qbPolish)
                        raiseEvent_("msgEvent", _
                                    "Removing Polish operation " & _
                                    objPolishHandle.ToString)
                        objPolishHandle.dispose()
                        .colPolish.Remove(intIndex1)
                        intIndex3 = intIndex3 + 1
                        raiseEvent_("codeRemoveEvent", intIndex1)
                    Next intIndex1
                End If
                intIndex1 = intIndex2 + CInt(IIf(intIndex3 = 0, 1, 0)) - intIndex3
                ' --- Extend permanent label table with new labels
                With colNewLabels
                    For intIndex2 = 1 To .Count
                        If OBJstate.usrState.booAssemblyRemovesCode Then
                            ' The position of the label was unknown
                            strKey = CStr(.Item(intIndex2))
                            intIndex3 = intIndex1
                        Else
                            ' The position of the label was known
                            colEntry = CType(.Item(intIndex2), Collection)
                            With colEntry
                                strKey = CStr(.Item(1)) : intIndex3 = CInt(.Item(2))
                            End With
                        End If
                        Try
                            colEntry = New Collection
                            With colEntry
                                .Add(strKey) : .Add(intIndex3)
                            End With
                        Catch objException As Exception
                            errorHandler_("Can't create label entry: " & _
                                          Err.Number & " " & Err.Description, _
                                          "assembler_", _
                                          "Marking object unusable and returning False", _
                                          objException)
                            OBJstate.usrState.booUsable = False : Return (False)
                        End Try
                        OBJstate.usrState.colLabel.Add(colEntry, strKey)
                    Next intIndex2
                    OBJcollectionUtilities.collectionClear(colNewLabels, booSetToNothing:=False)
                End With
            Loop
            ' --- Pass two: replace labels by addresses and remove comments
            For intIndex1 = 1 To .colPolish.Count
                If loopEventInterface_("Assembling code in pass two", _
                                       "Opcode", _
                                       intIndex1, _
                                       .colPolish.Count, _
                                       0, _
                                       "") Then Exit For
                objPolishHandle = CType(.colPolish.Item(intIndex1), qbPolish.qbPolish)
                With objPolishHandle
                    If _OBJqbOp.isJumpOp(.Opcode) Then
                        strKey = object2Key_(.Operand)
                        Try
                            .Operand = CType(OBJstate.usrState.colLabel.Item(strKey), Collection).Item(2)
                        Catch objException As Exception
                            If IsNumeric(.Operand) Then
                                errorHandler_("Undefined statement number " & CStr(.Operand), _
                                              "assembler_", _
                                              "Marking object unusable: returning False", _
                                              objException)
                            Else
                                errorHandler_("Internal compiler error: " & _
                                              "undefined label " & _
                                              _OBJutilities.enquote(CStr(.Operand)) & " " & _
                                              "at instruction " & intIndex1, _
                                              "assembler_", _
                                              "Marking object unusable: returning False", _
                                              objException)
                            End If
                            OBJstate.usrState.booUsable = False
                            Return (False)
                        End Try
                    End If
                    If OBJstate.usrState.booAssemblyRemovesCode Then
                        .Comment = ""
                        raiseEvent_("codeChangeEvent", objPolishHandle, intIndex1)
                    End If
                End With
            Next intIndex1
            .booAssembled = True
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Check the availavility of the object
    '
    '
    Private Function checkAvailability_(ByVal strProcedure As String, _
                                        ByVal strHelp As String) _
            As Boolean
        If Not OBJthreadStatus.available Then
            errorHandler_("Object instance is not available", _
                            strProcedure, _
                            strHelp, _
                            Nothing)
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Check the usability of the object
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String, _
                                  ByVal strHelp As String) _
            As Boolean
        With OBJstate.usrState
            If Not .booUsable Then
                errorHandler_("Object instance is not usable", _
                              strProcedure, _
                              strHelp, _
                              Nothing)
                Return (False)
            End If
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Clear the compiler object
    '
    '
    Private Function clear_() As Boolean
        With OBJstate.usrState
            Try
                If Not (.colLabel Is Nothing) Then
                    .objCollectionUtilities.collectionClear(.colLabel)
                End If
                If Not (.colPolish Is Nothing) Then
                    .objCollectionUtilities.collectionClear(.colVariables)
                End If
                If Not (.colReadData Is Nothing) Then
                    .objCollectionUtilities.collectionClear(.colReadData)
                End If
                If Not (.colSubFunctionIndex Is Nothing) Then
                    .objCollectionUtilities.collectionClear(.colSubFunctionIndex)
                End If
                If Not (.colVariables Is Nothing) Then
                    .objCollectionUtilities.collectionClear(.colVariables)
                End If
                If Not (.objScanner Is Nothing) Then
                    .objScanner.clear()
                    .objScanner = Nothing
                End If
            Catch objException As Exception
                errorHandler_("Could not clear scanner or parse tree: " & _
                              Err.Number & " " & Err.Description, _
                              "clear", "", _
                              objException)
                OBJstate.usrState.booUsable = False : Return (False)
            End Try
            .booAssemblyRemovesCode = True   ' Default value
        End With
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Clear the Polish collection
    '
    '
    ' --- Default syntax
    Private Overloads Function clearPolish_(ByRef colPolish As Collection) As Boolean
        Dim intIndex1 As Integer
        With colPolish
            For intIndex1 = 1 To .Count
                If loopEventInterface_("Disposing all Polish opcodes", _
                                       "opcode", _
                                       intIndex1, _
                                       .Count, _
                                       0, _
                                       "") Then Exit For
                CType(.Item(intIndex1), qbPolish.qbPolish).dispose()
            Next intIndex1
        End With
        Return (OBJstate.usrState.objCollectionUtilities.collectionClear(colPolish))
    End Function

    ' -----------------------------------------------------------------
    ' Clear the scanner object
    '
    '
    Private Function clearScanner_(ByRef objScanner As qbScanner.qbScanner) As Boolean
        If Not objScanner.dispose(OBJstate.usrState.booInspectCompilerObjects) Then
            Return (False)
        End If
        objScanner = Nothing
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Clear storage on behalf of clearStorage
    '
    '
    Private Function clearStorage_() As Boolean
        With OBJstate.usrState
            If (.colVariables Is Nothing) Then Return (True)
            With .colVariables
                Dim intIndex1 As Integer
                For intIndex1 = 1 To .Count
                    If loopEventInterface_("Clearing storage", _
                                                "variable", _
                                                intIndex1, _
                                                .Count, _
                                                0, _
                                                "") Then Exit For
                    CType(.Item(intIndex1), qbVariable.qbVariable).clearVariable()
                Next intIndex1
            End With
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Clear the variables collection
    '
    '
    Private Function clearVariables_(ByRef colVariables As Collection) As Boolean
        Dim intIndex1 As Integer
        With colVariables
            For intIndex1 = 1 To .Count
                If loopEventInterface_("Disposing variables", _
                                       "Variable", _
                                       intIndex1, _
                                       .Count, _
                                       0, _
                                       "") Then Exit For
                If Not disposeVariable_(CType(.Item(intIndex1), qbVariable.qbVariable)) Then
                    Return (False)
                End If
            Next intIndex1
        End With
        Return (OBJstate.usrState.objCollectionUtilities.collectionClear(colVariables))
    End Function

    ' -------------------------------------------------------------------------
    Private Function clone_() As qbQuickBasicEngine
        Return (CType(clone__(), qbQuickBasicEngine))
    End Function
    ' --- Private implementation  
    Private Function clone__() As Object Implements ICloneable.Clone
        With OBJstate.usrState
            Dim objClone As quickBasicEngine.qbQuickBasicEngine
            Try
                objClone = New quickBasicEngine.qbQuickBasicEngine
            Catch objException As Exception
                errorHandler_("Cannot create the clone object: " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              "Marking instance object as unusable: " & _
                              "returning False", _
                              objException)
                .booUsable = False
                Return (False)
            End Try
            objClone.ConstantFolding = .booConstantFolding
            objClone.DegenerateOpRemoval = .booDegenerateOpRemoval
            If Not (.objScanner Is Nothing) Then
                objClone.qbScanner = .objScanner.clone
            End If
            objClone.AssemblyRemovesCode = .booAssemblyRemovesCode
#If QUICKBASICENGINE_EXTENSION Then
            objClone.GenerateNOPs = .booGenerateNOPs
#End If
            Return (objClone)
        End With
    End Function

    ' -------------------------------------------------------------------------
    ' CompareTo implementation
    '
    '
    Private Function compareTo_(ByVal objQBE As Object) As Integer _
            Implements IComparable.CompareTo
        With OBJstate.usrState
            Dim objQBEhandle As quickBasicEngine.qbQuickBasicEngine = _
                CType(objQBE, quickBasicEngine.qbQuickBasicEngine)
            Dim booGenerateNOPs As Boolean = True
#If QUICKBASICENGINE_EXTENSION Then
            booGenerateNOPs = objQBEhandle.GenerateNOPs AndAlso .booGenerateNOPs
#End If
            Return (CInt(IIf(objQBEhandle.ConstantFolding = .booConstantFolding _
                            AndAlso _
                            objQBEhandle.DegenerateOpRemoval = .booDegenerateOpRemoval _
                            AndAlso _
                            (.objScanner Is Nothing AndAlso objQBEhandle.Scanner Is Nothing _
                             OrElse _
                             Not (.objScanner Is Nothing) _
                             AndAlso _
                             Not (objQBEhandle.Scanner Is Nothing) _
                             AndAlso _
                             objQBEhandle.qbScanner.compareTo(.objScanner)) _
                            AndAlso _
                            objQBEhandle.AssemblyRemovesCode = .booAssemblyRemovesCode _
                            AndAlso _
                            booGenerateNOPs, _
                            1, 0)))
        End With
    End Function


    ' -----------------------------------------------------------------
    ' Compile (private procedure, without thread logic)
    '
    '
    Private Function compile_() As Boolean
        With OBJstate.usrState
            Return (compiler_(.objScanner, _
                              .colPolish, _
                              .objScanner.SourceCode, _
                              .enuSourceCodeType, _
                              .colVariables, _
                              .objImmediateResult))
        End With
    End Function


#Region "Compiler"
    ' ----------------------------------------------------------------------
    ' Compile command, creating a reverse-Polish representation for execution
    '
    '
    ' This recursive-descent parser consists of a series of recognizers
    ' (names commencing with compiler__.)
    '
    ' Each recognizer is passed an index to the scanned command, the scan
    ' array, the Polish array, and the index of the end of the tokens to 
    ' be scanned (this last is needed when we're parsing an expression 
    ' in parentheses.)
    '
    ' Each recognizer attempts manfully to drive the parse forward by
    ' finding the goal it is given, and on success each recognizer sets
    ' its index parameter (passed by reference) to one past the last character
    ' of the substring that satisfies the goal, and returns True.  On failure
    ' each recognizer returns False.
    '
    '
    Private Function compiler_(ByVal objScanned As qbScanner.qbScanner, _
                               ByRef colPolish As Collection, _
                               ByVal strSourceCode As String, _
                               ByVal enuSourceType As ENUsourceCodeType, _
                               ByRef colVariables As Collection, _
                               ByRef objImmediateValue As qbVariable.qbVariable) As Boolean
        ' --- Parse the scanned tokens now
        Dim booOK As Boolean
        Dim intIndex1 As Integer = 1
        Dim objConstantValue As qbVariable.qbVariable
        Dim objDateTimeStart As New System.DateTime
        objDateTimeStart = Now
        Select Case enuSourceType
            Case ENUsourceCodeType.immediateCommand
                If Not compiler__immediateCommand_(intIndex1, _
                                                   objScanned, _
                                                   colPolish, _
                                                   objScanned.TokenCount, _
                                                   strSourceCode, _
                                                   colVariables, _
                                                   objImmediateValue, _
                                                   2) Then
                    Return (False)
                End If
            Case ENUsourceCodeType.program
                If Not compiler__sourceProgram_(intIndex1, _
                                                objScanned, _
                                                colPolish, _
                                                objScanned.TokenCount, _
                                                strSourceCode, _
                                                colVariables, _
                                                2) Then
                    Return (False)
                End If
            Case ENUsourceCodeType.unknown
                booOK = True
                Try
                    AddHandler Me.compileErrorEvent, AddressOf compileErrorEvent_handler_
                    booOK = compiler__immediateCommand_(intIndex1, _
                                                        objScanned, _
                                                        colPolish, _
                                                        objScanned.TokenCount, _
                                                        strSourceCode, _
                                                        colVariables, _
                                                        objImmediateValue, _
                                                        2)
                    RemoveHandler Me.compileErrorEvent, AddressOf compileErrorEvent_handler_
                Catch
                    booOK = False
                End Try
                If booOK AndAlso intIndex1 > objScanned.TokenCount Then
                    OBJstate.usrState.enuSourceCodeType = ENUsourceCodeType.immediateCommand
                Else
                    If Not resetCompiler_() Then Return (False)
                    intIndex1 = 1
                    If compiler__sourceProgram_(intIndex1, _
                                                objScanned, _
                                                colPolish, _
                                                objScanned.TokenCount, _
                                                strSourceCode, _
                                                colVariables, _
                                                2) Then
                        If intIndex1 > objScanned.TokenCount Then
                            OBJstate.usrState.enuSourceCodeType = ENUsourceCodeType.program
                        Else
                            compiler__errorHandler_("There is/are " & objScanned.TokenCount - intIndex1 & " token(s) " & _
                                                    "that can't be compiled, starting with this line", _
                                                    intIndex1, _
                                                    objScanned, _
                                                    strSourceCode, _
                                                    "Examine the code above the line")
                            Return (False)
                        End If
                    Else
                        Return (False)
                    End If
                End If
        End Select
        If OBJstate.usrState.booConstantFolding _
           AndAlso _
           Not (objImmediateValue Is Nothing) Then
            compiler__genCode_(OBJstate.usrState.colPolish, _
                               1, _
                               ENUop.opPushLiteral, _
                               objImmediateValue.value)
        End If
        compiler__genCode_(colPolish, _
                           intIndex1, _
                           ENUop.opEnd, _
                           Nothing, _
                           "Generated at end of code")
        compiler__parseEvent_("program", _
                              False, _
                              1, _
                              OBJstate.usrState.objScanner.TokenCount, _
                              1, _
                              OBJstate.usrState.colPolish.Count, _
                              "", _
                              1)
        If intIndex1 <= objScanned.TokenCount Then
            compiler__errorHandler_("Cannot parse beyond this point", _
                                    intIndex1, _
                                    objScanned, _
                                    strSourceCode, _
                                    "Examine the code before this point")
            Return (False)
        End If
        OBJstate.usrState.booCompiled = True
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' addFactor := mulFactor [addFactorRHS]
    '
    '
    ' Note: this method is the first compiler method in alpha sequence to
    ' implement the recognition of a grammar category, and, the rest of this
    ' comment header describes the coding standards for methods that
    ' recognize grammar categories.
    '
    ' Each of these methods is Private and has the name compiler__<gc>_ where
    ' <gc> is the gramar category name.
    '
    ' Each of these methods has the parameter signature shown below for 
    ' addFactor.
    '
    ' On start-up, each of these methods calls the raiseEvent_ method for
    ' the parseStartEvent for the grammar category. Then most shall save
    ' the current index (which will point to the next scanToken) in a save
    ' area (intIndex1) for backup.
    '
    ' On failure these methods each return False. However, they do not return
    ' False directly. Instead they return parseFail_(gc) where gc is the grammar
    ' category and the parseFail_ method returns False. parseFail_ is responsible
    ' for Raising the event parseFailEvent.
    '
    ' It is important to include Return parseFail_ at the end of each procedure
    ' if at the end of any procedure the parse has failed, for if this is not
    ' done, the procedure will "drop down" to the End Function statement, and
    ' (correctly) return False and (correctly) parse, but without accurately
    ' reflecting the parse in its events.
    '
    ' On success these methods should (1) call the parseEvent with the parameters
    ' shown in this method and (2) return True.
    '
    '
    Private Function compiler__addFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "addFactor")
        Dim intIndex1 As Integer = intIndex
        Dim intCountPrevious As Integer = colPolish.Count
        If Not compiler__mulFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("addFactor")
        End If
        Dim intIndex2 As Integer = intIndex
        If Not compiler__addFactorRHS_(intIndex, _
                                       objScanned, _
                                       colPolish, _
                                       intEndIndex, _
                                       strSourceCode, _
                                       objConstantValue, _
                                       colVariables, _
                                       intCountPrevious + 1, _
                                       booSideEffects, _
                                       intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("addFactor")
        compiler__parseEvent_("addFactor", _
                              False, _
                              intIndex1, _
                              intIndex - intIndex1, _
                              intCountPrevious, _
                              colPolish.Count - intCountPrevious, _
                              "", _
                              intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' addFactorRHS := mulOp mulFactor [addFactorRHS]
    '
    '
    Private Function compiler__addFactorRHS_(ByRef intIndex As Integer, _
                                             ByVal objScanned As qbScanner.qbScanner, _
                                             ByRef colPolish As Collection, _
                                             ByVal intEndIndex As Integer, _
                                             ByVal strSourceCode As String, _
                                             ByRef objConstantValue As qbVariable.qbVariable, _
                                             ByRef colVariables As Collection, _
                                             ByVal intCountPrevious As Integer, _
                                             ByRef booSideEffects As Boolean, _
                                             ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "addFactorRHS")
        If intIndex > intEndIndex Then Return compiler__parseFail_("addFactorRHS")
        Dim intIndex1 As Integer = intIndex
        Dim strMulOp As String
        If Not compiler__mulOp_(objScanned, _
                                intIndex, _
                                strSourceCode, _
                                strMulOp, _
                                intLevel + 1) Then Return compiler__parseFail_("addFactorRHS")
        Dim objConstantValueRHS As qbVariable.qbVariable
        Dim booSideEffectsLHS As Boolean = booSideEffects
        If Not compiler__mulFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValueRHS, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("addFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_(strMulOp, _
                                                  booSideEffectsLHS, _
                                                  objConstantValue, _
                                                  booSideEffects, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__addFactorRHS_(intIndex, _
                                       objScanned, _
                                       colPolish, _
                                       intEndIndex, _
                                       strSourceCode, _
                                       objConstantValue, _
                                       colVariables, _
                                       intCountPrevious + 1, _
                                       booSideEffects, _
                                       intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("addFactorRHS")
        compiler__parseEvent_("addFactorRHS", _
                              False, _
                              intIndex1, _
                              intIndex - intIndex1, _
                              intCountPrevious, _
                              colPolish.Count - intCountPrevious, _
                              "", _
                              intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' addOp := +|-
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression, and returns the operator in a reference parameter.
    '
    '
    Private Function compiler__addOp_(ByVal objScanned As qbScanner.qbScanner, _
                                      ByRef intIndex As Integer, _
                                      ByVal strSourceCode As String, _
                                      ByRef strOp As String, _
                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "addOp")
        strOp = Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length)
        Select Case strOp
            Case "+"
            Case "-"
            Case Else : Return compiler__parseFail_("addOp")
        End Select
        compiler__parseEvent_("addOp", _
                              False, _
                              intIndex, _
                              1, _
                              0, 0, _
                              "", _
                              intLevel)
        intIndex += 1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' andFactor := [ Not ] notFactor
    '
    '
    Private Function compiler__andFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "andFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim booNegate As Boolean = compiler__checkToken_(intIndex, _
                                                         objScanned, _
                                                         strSourceCode, _
                                                         intEndIndex, _
                                                         "NOT", _
                                                         intLevel + 1)
        If Not compiler__notFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("andFactor")
        End If
        If booNegate Then
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opNot, _
                                      Nothing, _
                                      "Not operator") Then Return compiler__parseFail_("andFactor")
        End If
        compiler__parseEvent_("andFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' andOp := And|AndAlso
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression.
    '
    '
    Private Function compiler__andOp_(ByVal objScanned As qbScanner.qbScanner, _
                                      ByRef intIndex As Integer, _
                                      ByVal strSourceCode As String, _
                                      ByRef strOp As String, _
                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "andOp")
        strOp = Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length)
        Select Case UCase(strOp)
            Case "AND"
            Case "ANDALSO"
            Case Else : Return compiler__parseFail_("andOp")
        End Select
        intIndex += 1
        compiler__parseEvent_("andOp", False, intIndex - 1, 1, 0, 0, "", intLevel)
        Return (True)
    End Function

    ' ---------------------------------------------------------------------
    ' asClause := As typeName
    '
    '
    Private Function compiler__asClause_(ByRef intIndex As Integer, _
                                         ByVal objScanned As qbScanner.qbScanner, _
                                         ByVal intEndIndex As Integer, _
                                         ByVal strSourceCode As String, _
                                         ByVal intMemoryReference As Integer, _
                                         ByRef colVariables As Collection, _
                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "asClause")
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, objScanned, strSourceCode, intEndIndex, "AS", intLevel) Then
            Return compiler__parseFail_("asClause")
        End If
        Dim strTypeName As String
        If Not compiler__asTypeName_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     intLevel, _
                                     strTypeName) Then
            intIndex = intIndex1 : Return compiler__parseFail_("asClause")
        End If
        Try
            Dim objOldVariable As qbVariable.qbVariable = CType(colVariables.Item(intMemoryReference), _
                                                                qbVariable.qbVariable)
            If objOldVariable.Dope.isArray Then
                strTypeName = "Array," & strTypeName & "," & objOldVariable.Dope.BoundList
            End If
            Dim objNewVariable As qbVariable.qbVariable = _
                New qbVariable.qbVariable(strTypeName)
            objNewVariable.VariableName = objOldVariable.VariableName
            objNewVariable.Tag = objOldVariable.Tag
            With colVariables
                .Remove(intMemoryReference) : .Add(objNewVariable, objNewVariable.VariableName, intMemoryReference)
            End With
            objOldVariable.dispose() : objOldVariable = Nothing
        Catch objException As Exception
            errorHandler_("Not able to replace variable's entry in symbol table: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__asClause_", _
                          "Marking object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False : Return compiler__parseFail_("asClause")
        End Try
        compiler__parseEvent_("asClause", _
                              False, _
                              intIndex1, _
                              intIndex - intIndex1, _
                              0, 0, _
                              "Type is " & strTypeName, _
                              intLevel)
        Return True
    End Function

    ' --------------------------------------------------------------------------------
    ' typeName := Boolean | Byte | Integer | Long | Single | Double | String | Variant
    '
    '
    Private Function compiler__asTypeName_(ByRef intIndex As Integer, _
                                           ByVal objScanned As qbScanner.qbScanner, _
                                           ByVal strSourceCode As String, _
                                           ByVal intEndIndex As Integer, _
                                           ByVal intLevel As Integer, _
                                           ByRef strTypeName As String) As Boolean
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                     intLevel) Then Return False
        strTypeName = objScanned.sourceMid(intIndex - 1, 1)
        If Not _OBJqbVariableType.isScalarType(strTypeName) Then
            intIndex = intIndex1 : Return False
        End If
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Binary operation to code with optimization including constant
    ' folding and the removal of degenerate operations (whoopee)
    '
    '
    Private Function compiler__binaryOpGen_(ByVal strOp As String, _
                                            ByVal booLHSsideEffect As Boolean, _
                                            ByRef objLHSconstantValue As qbVariable.qbVariable, _
                                            ByVal booRHSsideEffect As Boolean, _
                                            ByRef objRHSconstantValue As qbVariable.qbVariable, _
                                            ByRef colPolish As Collection, _
                                            ByVal intIndex As Integer) As qbVariable.qbVariable
        Dim booConstantFolding As Boolean
        Dim enuOpcode As qbOp.qbOp.ENUop = compiler__token2Op_(strOp)
        Dim objConstantValue As qbVariable.qbVariable
        Dim strComment As String = "Replace stack(top) by " & _
                                   _OBJqbOp.opCodeToString(enuOpcode) & _
                                   "(stack(top-1), stack(top))"
        booConstantFolding = OBJstate.usrState.booConstantFolding
        If Not booConstantFolding _
           OrElse _
           (objLHSconstantValue Is Nothing) AndAlso (objRHSconstantValue Is Nothing) Then
            Dim booRotate As Boolean = Not (objLHSconstantValue Is Nothing) And (objRHSconstantValue Is Nothing)
            If Not (objLHSconstantValue Is Nothing) Then
                If Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opPushLiteral, _
                                          objLHSconstantValue, _
                                          "Push LHS constant value") Then Return (Nothing)
                objLHSconstantValue = Nothing
            End If
            If Not (objRHSconstantValue Is Nothing) Then
                If Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opPushLiteral, _
                                          objRHSconstantValue, _
                                          "Push RHS constant value") Then Return (Nothing)
                objRHSconstantValue = Nothing
            End If
            If booRotate Then
                If Not compiler__genCode_(colPolish, _
                                         intIndex, _
                                         ENUop.opRotate, _
                                         1, _
                                         "Reorder left and right hand sides") Then Return (Nothing)
            End If
            If Not compiler__genCode_(colPolish, intIndex, _
                                        enuOpcode, _
                                        Nothing, _
                                        strComment) Then
                Return (Nothing)
            End If
        ElseIf (objLHSconstantValue Is Nothing) AndAlso Not (objRHSconstantValue Is Nothing) Then
            ' Right hand side is a constant 
            If Not booLHSsideEffect _
                AndAlso _
                compiler__binaryOpGen__checkDegeneracy_(enuOpcode, _
                                                        objLHSconstantValue, _
                                                        objRHSconstantValue, _
                                                        intIndex, _
                                                        colPolish, _
                                                        objConstantValue) Then
                Return (objConstantValue)
            Else
                If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, objRHSconstantValue, "Push constant RHS value") Then
                    Return (Nothing)
                End If
                If Not compiler__genCode_(colPolish, intIndex, _
                                            enuOpcode, _
                                            Nothing, _
                                            strComment) Then Return (Nothing)
            End If
        ElseIf Not (objLHSconstantValue Is Nothing) AndAlso (objRHSconstantValue Is Nothing) Then
            ' Left hand side is a constant
            If Not booRHSsideEffect _
                AndAlso _
                compiler__binaryOpGen__checkDegeneracy_(enuOpcode, _
                                                        objLHSconstantValue, _
                                                        objRHSconstantValue, _
                                                        intIndex, _
                                                        colPolish, _
                                                        objConstantValue) Then
                Return (objConstantValue)
            End If
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, objLHSconstantValue, "Push constant LHS value") Then
                Return (Nothing)
            End If
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "Reorder the operands") Then
                Return (Nothing)
            End If
            If Not compiler__genCode_(colPolish, intIndex, _
                                    enuOpcode, _
                                    Nothing, _
                                    strComment) Then Return (Nothing)
        Else
            Return compiler__constantEval_(strOp, objLHSconstantValue, objRHSconstantValue)
        End If
        Return (Nothing)
    End Function

    ' ----------------------------------------------------------------------
    ' Check for the opportunity to eliminate a degenerate op (whee) on
    ' behalf of compiler__binaryOpGen_
    '
    '
    ' Note that this method enforces the degeneracies that result in a 
    ' nonconstant value.  See the compiler's methods of the form
    ' compiler__xxxRHS_ for enforcement of the degeneracies that result
    ' in non-constant values, for by the time we get to this code, the
    ' code to create the non-constant value has been generated, and it
    ' would be inconvenient to remove it.
    '
    '
    Private Function compiler__binaryOpGen__checkDegeneracy_(ByVal enuOpCode As qbOp.qbOp.ENUop, _
                                                             ByVal objLHSvalue As qbVariable.qbVariable, _
                                                             ByVal objRHSvalue As qbVariable.qbVariable, _
                                                             ByVal intIndex As Integer, _
                                                             ByRef colPolish As Collection, _
                                                             ByRef objConstantValueOut As qbVariable.qbVariable) As Boolean
        If compiler__binaryOpGen__checkDegeneracy__(enuOpCode, _
                                                    objLHSvalue, _
                                                    objRHSvalue, _
                                                    intIndex, _
                                                    colPolish, _
                                                    objConstantValueOut) Then
            compiler__extendLastComment_(colPolish, _
                                            "Optimizer removed degenerate op " & _
                                            CStr(IIf(objLHSvalue Is Nothing, "v", objLHSvalue)) & "," & _
                                            CStr(IIf(objRHSvalue Is Nothing, "v", objRHSvalue)) & "," & _
                                            _OBJqbOp.opCodeToString(enuOpCode))
            Return (True)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Degenerate op eliminator
    '
    '                                                     
    Private Function compiler__binaryOpGen__checkDegeneracy__(ByVal enuOpCode As qbOp.qbOp.ENUop, _
                                                              ByVal objLHSvalue As qbVariable.qbVariable, _
                                                              ByVal objRHSvalue As qbVariable.qbVariable, _
                                                              ByVal intIndex As Integer, _
                                                              ByRef colPolish As Collection, _
                                                              ByRef objConstantValueOut As qbVariable.qbVariable) As Boolean
        If Not OBJstate.usrState.booDegenerateOpRemoval Then Return (False)
        Dim booLHSdblValue As Boolean = Not (objLHSvalue Is Nothing)
        Dim booRHSdblValue As Boolean = Not (objRHSvalue Is Nothing)
        Dim dblLHSvalue As Double
        Dim dblRHSvalue As Double
        Try
            dblLHSvalue = CDbl(objLHSvalue.value)
        Catch
            booLHSdblValue = False
        End Try
        Try
            dblRHSvalue = CDbl(objRHSvalue.value)
        Catch
            booRHSdblValue = False
        End Try
        ' "a+0 and 0+a are replaced by a"                
        If enuOpCode = ENUop.opAdd Then
            If Not booLHSdblValue AndAlso booRHSdblValue AndAlso dblRHSvalue = 0 Then Return (True)
            If booLHSdblValue AndAlso Not booRHSdblValue AndAlso dblLHSvalue = 0 Then Return (True)
        End If
        ' "a-0 is replaced by a: 0-a is replaced by -a"                
        If enuOpCode = ENUop.opSubtract Then
            If Not booLHSdblValue AndAlso booRHSdblValue AndAlso dblRHSvalue = 0 Then Return (True)
            If booLHSdblValue AndAlso Not booRHSdblValue AndAlso dblLHSvalue = 0 Then
                compiler__genCode_(colPolish, intIndex, ENUop.opNegate, Nothing, "Results from optimization of 0-a")
                Return (True)
            End If
        End If
        ' "a*1 and 1*a are replaced by a"                
        If enuOpCode = ENUop.opMultiply Then
            If Not booLHSdblValue AndAlso booRHSdblValue AndAlso dblRHSvalue = 1 Then Return (True)
            If booLHSdblValue AndAlso Not booRHSdblValue AndAlso dblLHSvalue = 1 Then Return (True)
        End If
        ' "a/1 is replaced by a"            
        If enuOpCode = ENUop.opDivide Then
            If Not booLHSdblValue AndAlso booRHSdblValue AndAlso dblRHSvalue = 1 Then Return (True)
        End If
        ' "a&Nil and Nil&a are replaced by a when Nil is the null string"   
        If enuOpCode = ENUop.opConcat Then
            Dim booLHSstrValue As Boolean = True
            Dim booRHSstrValue As Boolean = True
            Dim strLHSvalue As String
            Dim strRHSvalue As String
            Try
                strLHSvalue = CStr(objLHSvalue.value)
            Catch
                booLHSstrValue = False
            End Try
            Try
                strRHSvalue = CStr(objRHSvalue.value)
            Catch
                booRHSstrValue = False
            End Try
            If Not booLHSstrValue AndAlso booRHSstrValue AndAlso strRHSvalue = "" Then Return (True)
            If booLHSstrValue AndAlso Not booRHSstrValue AndAlso strLHSvalue = "" Then Return (True)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Check for a specific token, advancing index when found
    '
    '
    ' --- Check for a token type
    Private Overloads Function compiler__checkToken_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByVal strSourceCode As String, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal enuTokenTypeExpected As qbTokenType.qbTokenType.ENUtokenType, _
                                                     ByVal intLevel As Integer) As Boolean
        Return (compiler__checkToken__(intIndex, _
                                      objScanned, _
                                      strSourceCode, _
                                      intEndIndex, _
                                      enuTokenTypeExpected, _
                                      "", _
                                      intLevel))
    End Function
    ' --- Check for a token value
    Private Overloads Function compiler__checkToken_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByVal strSourceCode As String, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strTokenExpected As String, _
                                                     ByVal intLevel As Integer) As Boolean
        Return (compiler__checkToken__(intIndex, _
                                      objScanned, _
                                      strSourceCode, _
                                      intEndIndex, _
                                      Nothing, _
                                      strTokenExpected, _
                                      intLevel))
    End Function
    ' --- Check for a token type and value
    Private Overloads Function compiler__checkToken__(ByRef intIndex As Integer, _
                                                      ByVal objScanned As qbScanner.qbScanner, _
                                                      ByVal strSourceCode As String, _
                                                      ByVal intEndIndex As Integer, _
                                                      ByVal enuTokenTypeExpected As qbTokenType.qbTokenType.ENUtokenType, _
                                                      ByVal strTokenExpected As String, _
                                                      ByVal intLevel As Integer) As Boolean
        If intIndex > intEndIndex Then Return (False)
        With objScanned.QBToken(intIndex)
            If enuTokenTypeExpected <> Nothing _
               AndAlso _
               enuTokenTypeExpected.ToString <> .TokenType.ToString Then Return (False)
            Dim strTokenActual As String = _
                UCase(compiler__getToken_(intIndex, objScanned, strSourceCode))
            If strTokenExpected <> "" _
               AndAlso _
               UCase(strTokenExpected) <> strTokenActual Then Return (False)
            intIndex += 1
            compiler__parseEvent_(strTokenActual, _
                                    True, _
                                    intIndex - 1, _
                                    1, _
                                    0, 0, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' concatFactor := relFactor [concatFactorRHS]
    '
    '
    Private Function compiler__concatFactor_(ByRef intIndex As Integer, _
                                             ByVal objScanned As qbScanner.qbScanner, _
                                             ByRef colPolish As Collection, _
                                             ByVal intEndIndex As Integer, _
                                             ByVal strSourceCode As String, _
                                             ByRef objConstantValue As qbVariable.qbVariable, _
                                             ByRef colVariables As Collection, _
                                             ByRef booSideEffects As Boolean, _
                                             ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "concatFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__relFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("concatFactor")
        End If
        If intIndex > intEndIndex Then
            compiler__parseEvent_("concatFactor", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "No RHS", _
                                    intLevel)
            Return (True)
        Else
            Dim intIndex2 As Integer = intIndex
            If Not compiler__concatFactorRHS_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            booSideEffects, _
                                            intLevel + 1) _
                AndAlso _
                intIndex2 <> intIndex Then Return compiler__parseFail_("concatFactor")
        End If
        compiler__parseEvent_("concatFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "Has RHS", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' concatFactorRHS = relOp relFactor [concatFactorRHS]
    '
    '
    Private Function compiler__concatFactorRHS_(ByRef intIndex As Integer, _
                                                ByVal objScanned As qbScanner.qbScanner, _
                                                ByRef colPolish As Collection, _
                                                ByVal intEndIndex As Integer, _
                                                ByVal strSourceCode As String, _
                                                ByRef objConstantValue As qbVariable.qbVariable, _
                                                ByRef colVariables As Collection, _
                                                ByRef booSideEffects As Boolean, _
                                                ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "concatFactorRHS")
        If intIndex > intEndIndex Then Return compiler__parseFail_("concatFactorRHS")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim strRelOp As String
        If Not compiler__relOp_(objScanned, intIndex, strSourceCode, strRelOp, intLevel + 1) Then
            Return compiler__parseFail_("concatFactorRHS")
        End If
        Dim booSideEffectsRHS As Boolean
        Dim objConstantValueRHS As qbVariable.qbVariable
        If Not compiler__relFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValueRHS, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("concatFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_(strRelOp, _
                                                  booSideEffects, _
                                                  objConstantValue, _
                                                  booSideEffectsRHS, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__concatFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("concatFactorRHS")
        compiler__parseEvent_("concatFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Evaluate a constant operation
    '
    '
    ' --- Binary operation
    Private Overloads Function compiler__constantEval_(ByVal strOp As String, _
                                                       ByVal objOperand1 As qbVariable.qbVariable, _
                                                       ByVal objOperand2 As qbVariable.qbVariable) As qbVariable.qbVariable
        Return (compiler__constantEval__(strOp, True, objOperand1, objOperand2))
    End Function
    ' --- Unary operation
    Private Overloads Function compiler__constantEval_(ByVal strOp As String, _
                                                       ByVal objOperand As qbVariable.qbVariable) As qbVariable.qbVariable
        Return (compiler__constantEval__(strOp, True, objOperand, Nothing))
    End Function
    ' --- Common operation
    Private Overloads Function compiler__constantEval__(ByVal strOp As String, _
                                                        ByVal booBinary As Boolean, _
                                                        ByVal objOperand1 As qbVariable.qbVariable, _
                                                        ByVal objOperand2 As qbVariable.qbVariable) As qbVariable.qbVariable
        Dim strConstantExpression As String
        If booBinary Then
            strConstantExpression = compiler__constantEval__object2String_(objOperand1) & _
                                    strOp & _
                                    compiler__constantEval__object2String_(objOperand2)
        Else
            strConstantExpression = strOp & compiler__constantEval__object2String_(objOperand2)
        End If
        Dim objValue As qbVariable.qbVariable = _
            retrieveConstantCache_(strConstantExpression)
        If (objValue Is Nothing) Then
            Try
                objValue = New qbVariable.qbVariable
            Catch objException As Exception
                errorHandler_("Internal compiler error: cannot create qbVariable: " & _
                              Err.Number & " " & Err.Description, _
                              "compiler__constantEval__", _
                              "Marking object unusable and returning False", _
                              objException)
                OBJstate.usrState.booUsable = False
                Return (Nothing)
            End Try
            Dim strDummy As String
            objValue = evaluate_(strConstantExpression, False, strDummy)
            If Not (objValue Is Nothing) Then
                updateConstantCache_(strConstantExpression, objValue)
            End If
        End If
        Return (objValue)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert object to string value on behalf of compiler__constantEval__ 
    '
    '
    Private Function compiler__constantEval__object2String_(ByVal objValue As Object) _
            As String
        Dim objReturnValue As Object
        If (TypeOf objValue Is qbVariable.qbVariable) Then
            With CType(objValue, qbVariable.qbVariable)
                If Not .isScalar Then
                    errorHandler_("Programming error: value for constant evaluation isn't a scalar", _
                                  "compiler__constantEval__object2String_", _
                                  "Returning a quoted null string", _
                                  Nothing)
                    Return """"""
                End If
                objReturnValue = .value
            End With
        Else
            objReturnValue = objValue
        End If
        If (TypeOf objReturnValue Is String) Then
            Return _OBJutilities.enquote(CStr(objReturnValue))
        Else
            Return CStr(objReturnValue)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Erase the compiler's output Polish code
    '
    '
    Private Function compiler__erasePolish_(ByRef colPolish As Collection) As Boolean
        With colPolish
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                If Not CType(.Item(intIndex1), qbPolish.qbPolish).dispose Then
                    Exit For
                End If
            Next intIndex1
            If intIndex1 <= .Count Then Return (False)
            Return (OBJstate.usrState.objCollectionUtilities.collectionClear(colPolish))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Compiler error handler
    '
    '
    ' --- Call with no help
    Private Overloads Sub compiler__errorHandler_(ByVal strError As String, _
                                                    ByVal intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByVal strSource As String)
        compiler__errorHandler_(strError, intIndex, objScanned, strSource, "")
    End Sub
    ' --- Call with help
    Private Overloads Sub compiler__errorHandler_(ByVal strError As String, _
                                                  ByVal intIndex As Integer, _
                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                  ByVal strSource As String, _
                                                  ByVal strHelp As String)
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intLength As Integer
        With objScanned.QBToken(CInt(Math.Min(objScanned.TokenCount, intIndex)))
            raiseEvent_("compileErrorEvent", _
                        strError, _
                        .StartIndex, _
                        .Length, _
                        .LineNumber, _
                        strHelp, _
                        sourceIndex2Code_(intIndex, objScanned, strSource))
        End With
    End Sub

    ' ----------------------------------------------------------------------
    ' explicitAssignment := LET implicitAssignment
    '
    '
    Private Function compiler__explicitAssignment_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef colVariables As Collection, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "explicitAssignment")
        Dim intCountPrevious As Integer = 1
        Dim intIndex1 As Integer = intIndex
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "LET", _
                                 intLevel + 1) Then
            If Not compiler__implicitAssignment_(intIndex, _
                                                 objScanned, _
                                                 colPolish, _
                                                 intEndIndex, _
                                                 strSourceCode, _
                                                 colVariables, _
                                                 intLevel + 1) Then
                Return compiler__parseFail_("explicitAssignment")
            End If
            compiler__parseEvent_("explicitAssignment", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("explicitAssignment")
    End Function

    ' ----------------------------------------------------------------------
    ' expression := orFactor [ orOp expression ]        
    '
    '
    ' --- Does not return a constant value
    Private Overloads Function compiler__expression_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        Dim booSideEffects As Boolean
        Dim objConstantValue As qbVariable.qbVariable
        Return (compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValue, _
                                     colVariables, _
                                     booSideEffects, _
                                     intLevel))
    End Function
    ' --- Returns a constant value in a reference parameter
    Private Overloads Function compiler__expression_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef objConstantValue As qbVariable.qbVariable, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        Dim booSideEffects As Boolean
        Return (compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValue, _
                                     colVariables, _
                                     booSideEffects, _
                                     intLevel))
    End Function
    ' --- Does not return a constant value but indicates whether expression has side effects
    Private Overloads Function compiler__expression_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByRef booSideEffects As Boolean, _
                                                     ByVal intLevel As Integer) As Boolean
        Dim objConstantValue As qbVariable.qbVariable
        Return (compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValue, _
                                     colVariables, _
                                     booSideEffects, _
                                     intLevel))
    End Function
    ' --- Returns a constant value in a reference parameter and indicates side effects
    Private Overloads Function compiler__expression_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef objConstantValue As qbVariable.qbVariable, _
                                                     ByRef colVariables As Collection, _
                                                     ByRef booSideEffects As Boolean, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "expression")
        Dim booSuccess As Boolean
        Dim intCountPrevious As Integer
        Dim intIndex1 As Integer = intIndex
        If Not (objConstantValue Is Nothing) Then
            disposeVariable_(objConstantValue)
            objConstantValue = Nothing
        End If
        If compiler__orFactor_(intIndex, _
                               objScanned, _
                               colPolish, _
                               intEndIndex, _
                               strSourceCode, _
                               objConstantValue, _
                               colVariables, _
                               booSideEffects, _
                               intLevel + 1) Then
            Dim strOp As String
            If intIndex > objScanned.TokenCount _
               OrElse _
               Not compiler__orOp_(objScanned, _
                                   intIndex, _
                                   strSourceCode, _
                                   strOp, _
                                   intLevel + 1) Then
                compiler__parseEvent_("expression", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "No Or occurs", _
                                        intLevel)
                Return (True)
            End If
            Dim strLbl As String
            If UCase(strOp) = "ORELSE" Then
                strLbl = compiler__getLabel_()
                If Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opDuplicate, _
                                          strLbl, _
                                          "OrElse: duplicate stack and skip RHS when LHS is True") _
                   OrElse _
                   Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opJumpNZ, _
                                          strLbl, _
                                          "") Then
                    Return compiler__parseFail_("expression")
                End If
            End If
            Dim booSideEffectsRHS As Boolean
            Dim objConstantValueRHS As qbVariable.qbVariable
            If compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValueRHS, _
                                     colVariables, _
                                     booSideEffectsRHS, _
                                     intLevel + 1) Then
                If UCase(strOp) = "ORELSE" Then
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLbl, "OrElse jump target for True") Then
                        Return compiler__parseFail_("expression")
                    End If
                Else
                    objConstantValue = compiler__binaryOpGen_("OR", _
                                                              booSideEffects, _
                                                              objConstantValue, _
                                                              booSideEffectsRHS, _
                                                              objConstantValueRHS, _
                                                              colPolish, _
                                                              intIndex)
                End If
                compiler__parseEvent_("expression", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "Or occurs", _
                                        intLevel)
                Return (True)
            End If
        End If
        Return compiler__parseFail_("expression")
    End Function

    ' ----------------------------------------------------------------------
    ' expressionList = expression [ comma expressionList ]
    '
    '
    ' This method parses four types of expression lists using three overloads.
    '
    '
    '      *  Dim statement expression lists which define dimensions
    '         and upper bounds (usrVariableDope parameter passed with
    '         undefined values that are assigned by a parse of the
    '         expressions in the list...which need to be constant.)
    '
    '      *  Function parameter lists (intMinParmsExpected and 
    '         intMaxParmsExpected parameters passed: reference
    '         parameter intParms set to actual argcount: intMaxParmsExpected
    '         may be -1 for no upper limit.)
    '
    '      *  Print statement lists (no optional parameters)
    '
    '      *  Subscript expression lists (usrVariableDope parameter passed
    '         with predefined values which are reconciled with the count
    '         of list elements and the constant values in the list.)  In
    '         this mode, the intMemoryReference parameter is assigned to
    '         -1 or the constant value specified by the expressions in a list
    '         of constant subscripts.
    '
    '
    ' --- Dim statement/lValue call
    Private Overloads Function compiler__expressionList_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByRef intMemoryReference As Integer, _
                                                         ByRef colVariables As Collection, _
                                                         ByVal intSymIndex As Integer, _
                                                         ByVal intLevel As Integer) As Boolean
        Dim intParms As Integer
        Return (compiler__expressionList_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         CType(IIf(CType(colVariables.Item(intSymIndex), _
                                                         qbVariable.qbVariable).Dope.isArray, _
                                                   ENUexpressionListCallMode.subscriptList, _
                                                   ENUexpressionListCallMode.dimStatement), _
                                               ENUexpressionListCallMode), _
                                         intMemoryReference, _
                                         0, _
                                         0, _
                                         colVariables, _
                                         intSymIndex, _
                                         intParms, _
                                         intLevel))
    End Function
    ' --- Function call
    Private Overloads Function compiler__expressionList_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByVal intMinParmsExpected As Integer, _
                                                         ByVal intMaxParmsExpected As Integer, _
                                                         ByRef colVariables As Collection, _
                                                         ByRef intParms As Integer, _
                                                         ByVal intLevel As Integer) As Boolean
        Return (compiler__expressionList_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         ENUexpressionListCallMode.functionCall, _
                                         Nothing, _
                                         intMinParmsExpected, _
                                         intMaxParmsExpected, _
                                         colVariables, _
                                         0, _
                                         intParms, _
                                         intLevel))
    End Function
    ' --- Print statement
    Private Overloads Function compiler__expressionList_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByRef colVariables As Collection, _
                                                         ByVal intLevel As Integer) As Boolean
        Dim intParms As Integer
        Return (compiler__expressionList_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         ENUexpressionListCallMode.printStatement, _
                                         Nothing, _
                                         0, _
                                         0, _
                                         colVariables, _
                                         0, _
                                         intParms, _
                                         intLevel))
    End Function
    ' --- Common logic
    Private Enum ENUexpressionListCallMode
        dimStatement
        functionCall
        printStatement
        subscriptList
    End Enum
    Private Overloads Function compiler__expressionList_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByVal enuCallMode As ENUexpressionListCallMode, _
                                                         ByRef intMemoryReference As Integer, _
                                                         ByVal intMinParmsExpected As Integer, _
                                                         ByVal intMaxParmsExpected As Integer, _
                                                         ByRef colVariables As Collection, _
                                                         ByVal intSymIndex As Integer, _
                                                         ByRef intParms As Integer, _
                                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "expressionList")
        Dim booSaveConstantFolding As Boolean
        Dim booSideEffects As Boolean
        Dim booTOclause As Boolean
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intExpressionCount As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer = intIndex
        Dim intIndex3 As Integer
        Dim intLBound As Integer
        Dim intStartIndex As Integer
        Dim intUBound As Integer
        Dim objConstantValue As qbVariable.qbVariable
        Dim objConstantValue2 As qbVariable.qbVariable
        Dim objMemoryReference As Object
        Dim strSubscriptList As String
        intExpressionCount = 0
        If enuCallMode = ENUexpressionListCallMode.subscriptList Then
            objMemoryReference = intSymIndex
        End If
        intParms = 0
        Do
            intStartIndex = intIndex
            intIndex1 = colPolish.Count
            If enuCallMode <> ENUexpressionListCallMode.printStatement Then
                booSaveConstantFolding = OBJstate.usrState.booConstantFolding
                OBJstate.usrState.booConstantFolding = True
            End If
            Dim booOK As Boolean = compiler__expression_(intIndex, _
                                                         objScanned, _
                                                         colPolish, _
                                                         intEndIndex, _
                                                         strSourceCode, _
                                                         objConstantValue, _
                                                         colVariables, _
                                                         intLevel + 1)
            If enuCallMode <> ENUexpressionListCallMode.printStatement Then
                OBJstate.usrState.booConstantFolding = booSaveConstantFolding
            End If
            If booOK Then
                intExpressionCount += 1
                If enuCallMode = ENUexpressionListCallMode.functionCall Then
                    If intMaxParmsExpected <> -1 AndAlso intExpressionCount > intMaxParmsExpected Then
                        compiler__errorHandler_("In a function call, there are too many parameters.  The maximum is " & _
                                                intMaxParmsExpected, _
                                                intIndex, _
                                                objScanned, _
                                                strSourceCode, _
                                                "This function does not accept these parameters")
                        Return compiler__parseFail_("expressionList")
                    End If
                    If Not (objConstantValue Is Nothing) Then
                        If Not compiler__genCode_(colPolish, intIndex, _
                                                  ENUop.opPushLiteral, _
                                                  objConstantValue, _
                                                  "Function parameter " & intExpressionCount) Then
                            Return compiler__parseFail_("expressionList")
                        End If
                    End If
                Else
                    If enuCallMode = ENUexpressionListCallMode.dimStatement Then
                        With CType(colVariables.Item(intSymIndex), qbVariable.qbVariable)
                            booTOclause = False
                            If compiler__checkToken_(intIndex, _
                                                     objScanned, _
                                                     strSourceCode, _
                                                     intEndIndex, _
                                                     "TO", _
                                                     intLevel + 1) Then
                                booTOclause = True
                                intIndex3 = objScanned.findRightParenthesis(intIndex, _
                                                                            intEndIndex:=intEndIndex)
                                If Not compiler__expression_(intIndex, _
                                                              objScanned, _
                                                             colPolish, _
                                                             intIndex2, _
                                                             strSourceCode, _
                                                             objConstantValue2, _
                                                             colVariables, _
                                                             intLevel + 1) Then
                                    compiler__errorHandler_("TO in array dimension is not followed " & _
                                                            "by a constant expression", _
                                                            intIndex, _
                                                            objScanned, _
                                                            strSourceCode, _
                                                            "The syntax should be lowerBound To upperBound")
                                    Return compiler__parseFail_("expressionList")
                                End If
                            End If
                            If (objConstantValue Is Nothing) _
                               OrElse _
                               booTOclause _
                               AndAlso _
                               (objConstantValue2 Is Nothing) Then
                                compiler__errorHandler_("Dim statement doesn't use constants exclusively", _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode, _
                                                        "No Dim statement can use variables this way")
                            End If
                            Try
                                If booTOclause Then
                                    intLBound = CInt(objConstantValue.value)
                                    intUBound = CInt(objConstantValue2.value)
                                Else
                                    intLBound = 0 : intUBound = CInt(objConstantValue.value)
                                End If
                            Catch : End Try
                            .fromString(CStr(IIf(.Dope.isArray, .Dope.ToString, "Array,Variant")) & _
                                        "," & _
                                        CStr(intLBound) & "," & CStr(intUBound))
                        End With
                    ElseIf enuCallMode = ENUexpressionListCallMode.subscriptList Then
                       With CType(colVariables.Item(intSymIndex), qbVariable.qbVariable).Dope
                            If Not .isArray Then
                                compiler__errorHandler_("Subscript used in nonarray", _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode)
                                Return compiler__parseFail_("expressionList")
                            End If
                            If intExpressionCount > .Dimensions Then
                                compiler__errorHandler_("Too many subscripts for an array defined as having " & _
                                                        .Dimensions & " dimensions", _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode)
                                Return compiler__parseFail_("expressionList")
                            End If
                            objMemoryReference = Nothing
                            Dim intSubscript As Integer = _
                                compiler__expressionList_object2IntegerSub_(objConstantValue, _
                                                                            False, _
                                                                            intIndex, _
                                                                            objScanned, _
                                                                            strSourceCode)
                            If Not (objConstantValue Is Nothing) Then
                                compiler__genCode_(colPolish, _
                                                   intIndex, _
                                                   ENUop.opPushLiteral, _
                                                   intSubscript, _
                                                   "Push constant subscript for dimension " & _
                                                   intExpressionCount)
                            Else
                                compiler__genCode_(colPolish, _
                                                   intIndex, _
                                                   ENUop.opPush, _
                                                   intSubscript, _
                                                   "Variable subscript provided for dimension " & _
                                                   intExpressionCount)
                            End If
                        End With
                    ElseIf enuCallMode = ENUexpressionListCallMode.printStatement Then
                        If Not (objConstantValue Is Nothing) Then
                            If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, _
                                                        objConstantValue, _
                                                        "Push constant value") Then
                                Return compiler__parseFail_("expressionList")
                            End If
                        End If
                    Else
                        errorHandler_("Internal compiler error: invalid enuCallMode parameter " & _
                                      "passed to compiler__expressionList_ has value " & enuCallMode, _
                                      "compiler__expressionList_", _
                                      "Making object unusable and returning False")
                        Return compiler__parseFail_("expressionList")
                    End If
                End If
            Else
                Return compiler__parseFail_("expressionList")
            End If
            intParms += 1
            If Not (compiler__checkToken_(intIndex, _
                                          objScanned, _
                                          strSourceCode, _
                                          intEndIndex, _
                                          qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma, _
                                          intLevel + 1) _
                    OrElse _
                    enuCallMode = ENUexpressionListCallMode.printStatement _
                    AndAlso _
                    compiler__checkToken_(intIndex, _
                                          objScanned, _
                                          strSourceCode, _
                                          intEndIndex, _
                                          qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon, _
                                          intLevel + 1)) Then
                Exit Do
            End If
        Loop
        If enuCallMode = ENUexpressionListCallMode.functionCall Then
            If intExpressionCount < intMinParmsExpected Then
                compiler__errorHandler_("In function call, there are too few parameters.  The Math.min is " & _
                                        intMinParmsExpected, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "")
                Return compiler__parseFail_("expressionList")
            End If
            If intMaxParmsExpected <> -1 AndAlso intExpressionCount > intMaxParmsExpected Then
                compiler__errorHandler_("In function call, there are too many parameters.  The Math.max is " & _
                                        intMaxParmsExpected, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "")
                Return compiler__parseFail_("expressionList")
            End If
        End If
        If enuCallMode = ENUexpressionListCallMode.subscriptList Then
            ' Create subscript stack frame: array, count: it will be followed 
            ' by the subscripts
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opPushLiteral, _
                                      intExpressionCount, _
                                      "Push subscript count") Then
                Return compiler__parseFail_("expressionList")
            End If
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opPush, _
                                      intSymIndex, _
                                      "Push array variable's address") Then
                Return compiler__parseFail_("expressionList")
            End If
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opPushIndirect, _
                                      Nothing, _
                                      "Push the array variable") Then
                Return compiler__parseFail_("expressionList")
            End If
        End If
        If enuCallMode = ENUexpressionListCallMode.subscriptList Then
            Dim objHandle As qbVariable.qbVariable = CType(colVariables.Item(intSymIndex), _
                                                           qbVariable.qbVariable)
            If objHandle.Dope.Dimensions _
               <> _
               intParms Then
                compiler__errorHandler_("Wrong number of subscripts provided for the array " & _
                                        objHandle.VariableName, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return False
            End If
        End If
        compiler__parseEvent_("expressionList", _
                                False, _
                                intIndex2, _
                                intIndex - intIndex2, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Extend last comment
    '
    '
    Private Sub compiler__extendLastComment_(ByRef colPolish As Collection, ByVal strAdd As String)
        Dim intUBound As Integer = -1
        Try
            intUBound = colPolish.Count
        Catch
            Return
        End Try
        _OBJutilities.append(CType(OBJstate.usrState.colPolish.Item(intUBound), _
                                   qbPolish.qbPolish).Comment, _
                             ": ", _
                             strAdd)
    End Sub

    ' ----------------------------------------------------------------------
    ' Locate new line (new line or colon)
    '
    '
    ' Returns one beyond the end of the source when the newline cannot be
    ' found.
    '
    ' See also compiler__logicalNewline_ for the method that increments the
    ' scan index when a newline or colon only is found.  The purpose
    ' of this method is to locate newline for Data and Print.
    '
    '
    Private Function compiler__findNewline_(ByVal intIndex As Integer, _
                                            ByVal objScanned As qbScanner.qbScanner, _
                                            ByVal intEndIndex As Integer, _
                                            ByVal booPrintStatement As Boolean, _
                                            Optional ByVal booColonIsNewLine As Boolean = True) As Integer
        Dim intIndex1 As Integer
        For intIndex1 = intIndex To intEndIndex
            With objScanned.QBToken(intIndex1)
                If .TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline _
                   OrElse _
                   booColonIsNewLine _
                   AndAlso _
                   .TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon _
                   OrElse _
                   booPrintStatement _
                   AndAlso _
                   intIndex1 < intEndIndex _
                   AndAlso _
                   objScanned.QBToken(intIndex1 + 1).TokenType _
                   = _
                   qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline _
                   AndAlso _
                   .TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon Then
                    Exit For
                End If
            End With
        Next intIndex1
        Return (intIndex1)
    End Function

    ' ----------------------------------------------------------------------
    ' Lookup one variable: if it cannot be found, add it automatically
    ' unless Option Explicit is in effect
    '
    '
    Private Function compiler__findOrAddVariable_(ByVal strVariable As String, _
                                                  ByRef intIndex As Integer, _
                                                  ByRef colVariables As Collection, _
                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                  ByVal strSourceCode As String, _
                                                  ByVal booDefContext As Boolean) As Boolean
        With OBJstate.usrState
            Dim objVariable As qbVariable.qbVariable
            Try
                objVariable = CType(colVariables.Item(strVariable), qbVariable.qbVariable)
            Catch : End Try
            If (objVariable Is Nothing) Then
                If booDefContext OrElse Not .booExplicit Then
                    Try
                        Dim objNew As qbVariable.qbVariable
                        objNew = New qbVariable.qbVariable("Variant")
                        With objNew
                            .Tag = colVariables.Count + 1
                            .VariableName = strVariable
                        End With
                        colVariables.Add(objNew, strVariable)
                    Catch objException As Exception
                        errorHandler_("Internal compiler error: cannot expand storage: " & _
                                      Err.Number & " " & Err.Description, _
                                      "compiler__findOrAddVariable_", _
                                      "Making object unusable and returning False", _
                                      objException)
                        OBJstate.usrState.booUsable = False
                        Return (False)
                    End Try
                    intIndex = colVariables.Count
                Else
                    compiler__errorHandler_("Undefined variable " & _
                                            _OBJutilities.enquote(strVariable) & ": " & _
                                            "Option Explicit is in effect", _
                                            intIndex, _
                                            objScanned, _
                                            strSourceCode, _
                                            "")
                    Return (False)
                End If
            Else
                intIndex = CInt(objVariable.Tag)
            End If
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' formalParameterDef := [ ByVal | ByRef ] identifier ["()"] asClause
    '
    '
    ' Note: formal parameters are not implemented yet
    '
    '
    Private Function compiler__formalParameterDef_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef colVariables As Collection, _
                                                   ByRef objNestingStack As Stack, _
                                                   ByRef usrFormalParameters() As TYPformalParameters, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "formalParameterDef")
        errorHandler_("Formal parameter is in use in source code but this feature hasn't " & _
                      "been implemented yet", _
                      "formalParameterDef", _
                      "Returning parse failure (False)", _
                      Nothing)
        Return compiler__parseFail_("formalParameterDef")
    End Function

    ' ----------------------------------------------------------------------
    ' formalParameterList := ( formalParameterListBody )
    '
    '
    Private Function compiler__formalParameterList_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByRef colVariables As Collection, _
                                                    ByRef objNestingStack As Stack, _
                                                    ByRef usrFormalParameters() As TYPformalParameters, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "formalParameterList")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "(", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("formalParameterList")
        End If
        If Not compiler__formalParameterListBody_(intIndex, _
                                                  objScanned, _
                                                  colPolish, _
                                                  intEndIndex, _
                                                  strSourceCode, _
                                                  colVariables, _
                                                  objNestingStack, _
                                                  OBJstate.usrState.usrSubFunction(UBound(OBJstate.usrState.usrSubFunction)).usrFormalParameters, _
                                                  intLevel + 1) Then
            intIndex = intIndex1
            Return compiler__parseFail_("formalParameterList")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     ")", _
                                     intLevel + 1) Then
            intIndex = intIndex1
            Return compiler__parseFail_("formalParameterList")
        End If
        compiler__parseEvent_("formalParameterList", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' formalParameterListBody := formalParameterDef [ , formalParameterListBody ]
    '
    '
    Private Function compiler__formalParameterListBody_(ByRef intIndex As Integer, _
                                                        ByVal objScanned As qbScanner.qbScanner, _
                                                        ByRef colPolish As Collection, _
                                                        ByVal intEndIndex As Integer, _
                                                        ByVal strSourceCode As String, _
                                                        ByRef colVariables As Collection, _
                                                        ByRef objNestingStack As Stack, _
                                                        ByRef usrFormalParameters() As TYPformalParameters, _
                                                        ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "formalParameterListBody")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__formalParameterDef_(intIndex, _
                                             objScanned, _
                                             colPolish, _
                                             intEndIndex, _
                                             strSourceCode, _
                                             colVariables, _
                                             objNestingStack, _
                                             usrFormalParameters, _
                                             intLevel + 1) Then
            Return compiler__parseFail_("formalParameterListBody")
        End If
        Do
            If Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma, _
                                         intLevel + 1) Then Exit Do
            If Not compiler__formalParameterDef_(intIndex, _
                                                 objScanned, _
                                                 colPolish, _
                                                 intEndIndex, _
                                                 strSourceCode, _
                                                 colVariables, _
                                                 objNestingStack, _
                                                 usrFormalParameters, _
                                                 intLevel + 1) Then
                intIndex = intIndex1
                Return compiler__parseFail_("formalParameterListBody")
            End If
        Loop
        compiler__parseEvent_("formalParameterListBody", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
        compiler__parseFail_("formalParameterListBody")
    End Function

    ' ----------------------------------------------------------------------
    ' functionCall := functionName "(" expressionList ")"
    ' functionName := Abs | Asc | Ceil | Chr | Cos | Eval | Floor | Int | 
    '           Iif | Isnumeric | Lbound | Lcase | Left | Len | Log | 
    '           Max | Min | Mid | Replace | Right | Rnd | Sin | 
    '           Sgn | String | Tab | Trim |
    '           Ubound | Ucase | Utility
    '
    '
    ' --- Basic call
    Private Overloads Function compiler__functionCall_(ByRef intIndex As Integer, _
                                                       ByVal objScanned As qbScanner.qbScanner, _
                                                       ByRef colPolish As Collection, _
                                                       ByVal intEndIndex As Integer, _
                                                       ByVal strSourceCode As String, _
                                                       ByRef colVariables As Collection, _
                                                       ByVal intLevel As Integer) As Boolean
        Dim booSideEffects As Boolean
        compiler__functionCall_(intIndex, _
                                objScanned, _
                                colPolish, _
                                intEndIndex, _
                                strSourceCode, _
                                colVariables, _
                                booSideEffects, _
                                intLevel)
    End Function
    ' --- Tell caller about any side effects                                      
    Private Overloads Function compiler__functionCall_(ByRef intIndex As Integer, _
                                                       ByVal objScanned As qbScanner.qbScanner, _
                                                       ByRef colPolish As Collection, _
                                                       ByVal intEndIndex As Integer, _
                                                       ByVal strSourceCode As String, _
                                                       ByRef colVariables As Collection, _
                                                       ByRef booSideEffects As Boolean, _
                                                       ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "functionCall")
        raiseEvent_("parseStartEvent", "functionCall")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("functionCall")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "(", _
                                     intLevel + 1) Then
            intIndex -= 1
            Return compiler__parseFail_("functionCall")
        End If
        Dim intArgEnd As Integer = objScanned.findRightParenthesis(intIndex, _
                                                                   intEndIndex:=intEndIndex)
        If intArgEnd > intEndIndex OrElse compiler__getToken_(intArgEnd, objScanned, strSourceCode) <> ")" Then
            compiler__errorHandler_("Function call is followed by parenthesized argument list with unbalanced parentheses", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode, _
                                    "")
            Return compiler__parseFail_("functionCall")
        End If
        Dim strFunction As String
        Dim intUBound As Integer
        Dim intParms As Integer
        With objScanned.QBToken(intIndex - 2)
            strFunction = UCase(Mid(strSourceCode, .StartIndex, .Length))
            If Mid(strFunction, Len(strFunction)) = "$" Then strFunction = Mid(strFunction, 1, Len(strFunction) - 1)
            Select Case strFunction
                Case "ABS"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    Dim strLabel As String = compiler__getLabel_()
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opDuplicate, Nothing, "Get the absolute value: dup stack top") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, 0, "Push zero") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushGE, 0, "Compare values") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opJumpNZ, strLabel, "Jump when stack(t-1)>=0") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opNegate, 0, "Negative value") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "abs(x) is at top") Then Return compiler__parseFail_("functionCall")
                Case "ASC"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opAsc, 0, "Get ASCII code for character") Then Return compiler__parseFail_("functionCall")
                Case "CEIL"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opCeil, 0, "Get first int n > stack(top)") Then Return compiler__parseFail_("functionCall")
                Case "CHR"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opChr, 0, "Get character for ASCII code") Then Return compiler__parseFail_("functionCall")
                Case "COS"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opCos, 0, "Cosine") Then Return compiler__parseFail_("functionCall")
                Case "EVAL"
                    Dim intStartIndex As Integer = intIndex
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opEval, _
                                              0, _
                                              "Evaluate the Quickbasic expression (lightweight)" & _
                                              _OBJutilities.enquote(rebuildCode_(objScanned, _
                                                                                strSourceCode, _
                                                                                intStartIndex, _
                                                                                intIndex - 1))) Then Return compiler__parseFail_("functionCall")
                Case "EVALUATE"
                    Dim intStartIndex As Integer = intIndex
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opEvaluate, _
                                              0, _
                                              "Evaluate the Quickbasic expression (heavyweight)" & _
                                              _OBJutilities.enquote(rebuildCode_(objScanned, _
                                                                                strSourceCode, _
                                                                                intStartIndex, _
                                                                                intIndex - 1))) Then Return compiler__parseFail_("functionCall")
                Case "FLOOR"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opFloor, 0, "Get first int n < stack(top)") Then Return compiler__parseFail_("functionCall")
                Case "INT"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opInt, Nothing, "Round to integer function") Then Return compiler__parseFail_("functionCall")
                Case "IIF"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     3, _
                                                     3, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opIif, 0, "Conditional expression") Then Return compiler__parseFail_("functionCall")
                Case "ISNUMERIC"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opIsNumeric, 0, "Stack True when a stack(top) is a number, False otherwise") Then Return compiler__parseFail_("functionCall")
                Case "LBOUND"
                    If Not compiler__functionCall__bounds_(intIndex, _
                                                           objScanned, _
                                                           colPolish, _
                                                           intArgEnd, _
                                                           strSourceCode, _
                                                           False, _
                                                           colVariables, _
                                                           intLevel + 1) Then Return compiler__parseFail_("functionCall")
                Case "LCASE"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opLCase, 0, "Replace string by lower case") Then Return compiler__parseFail_("functionCall")
                Case "LEFT"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opMid, 0, "Execute Left as a Mid") Then Return compiler__parseFail_("functionCall")
                Case "LEN"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opLen, 0, "Get length of string") Then Return compiler__parseFail_("functionCall")
                Case "LOG"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opLog, 0, "Natural logarithm") Then Return compiler__parseFail_("functionCall")
                Case "MAX"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opMax, 0, "Maximum value") Then Return compiler__parseFail_("functionCall")
                Case "MID"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     3, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If intParms = 2 Then
                        If Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "Get string length: string,index->index,string") _
                           OrElse _
                           Not compiler__genCode_(colPolish, intIndex, ENUop.opDuplicate, Nothing, "index,string->index,string,string") _
                           OrElse _
                           Not compiler__genCode_(colPolish, intIndex, ENUop.opLen, 0, "index,string,string->index,string,length") _
                           OrElse _
                           Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 2, "index,string,length->length,string,index") _
                           OrElse _
                           Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "length,string,index->length,index,string") _
                           OrElse _
                           Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 2, "length,index,string->string,index,length") Then
                        End If
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opMid, 0, "Substring") Then Return compiler__parseFail_("functionCall")
                Case "MIN"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opMin, 0, "Math.Mininum value") Then Return compiler__parseFail_("functionCall")
                Case "REPLACE"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     3, _
                                                     3, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opReplace, 0, "Translates string1 to string2") Then Return compiler__parseFail_("functionCall")
                Case "RIGHT"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "Execute a RIGHT op: string,rightLength->rightLength,string") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opDuplicate, Nothing, "rightLength,string->rightLength,string,string") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opLen, 0, "rightLength,string,string->rightLength,string,stringLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "rightLength,string,stringLength->rightLength,stringLength,string") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 2, "rightLength,stringLength,string->string,stringLength,rightLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opDuplicate, Nothing, "string,stringLength,rightLength->string,stringLength,rightLength,rightLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 2, "string,stringLength,rightLength,rightLength->string,rightLength,rightLength,stringLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "string,rightLength,rightLength,stringLength->string,rightLength,stringLength,rightLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opSubtract, 0, "string,rightLength,stringLength,rightLength->string,rightLength,leftLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, 1, "string,rightLength,leftLength->string,rightLength,leftLength,1") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opAdd, 0, "string,rightLength,leftLength,1->string,rightLength,rightStart") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opRotate, 1, "string,rightLength,rightStart->string,rightStart,rightLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, 0, "string,rightStart,rightLength->string,rightStart,rightLength,0") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opMax, 1, "string,rightStart,rightLength,0->string,rightStart,rightLength") _
                            OrElse _
                            Not compiler__genCode_(colPolish, intIndex, ENUop.opMid, 0, "string,rightStart,rightLength->rightValue") Then Return compiler__parseFail_("functionCall")
                Case "RND"
                    intUBound = colPolish.Count
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     0, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              CType(IIf(intUBound <> colPolish.Count, _
                                                        ENUop.opRndSeed, _
                                                        ENUop.opRnd), _
                                                    qbOp.qbOp.ENUop), _
                                              0, _
                                              "Random number function") Then Return compiler__parseFail_("functionCall")
                Case "ROUND"
                    intUBound = colPolish.Count
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              ENUop.opRound, _
                                              0, _
                                              "Round double to integer fractional digits") Then Return compiler__parseFail_("functionCall")
                Case "SIN"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opSin, 0, "Sine") Then Return compiler__parseFail_("functionCall")
                Case "SGN"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opSgn, 0, "Signum") Then Return compiler__parseFail_("functionCall")
                Case "SQR"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opSqr, 0, "Square root") Then Return compiler__parseFail_("functionCall")
                Case "STRING"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     2, _
                                                     2, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opString, _
                                              0, _
                                              "Make copies of a character") Then Return compiler__parseFail_("functionCall")
                Case "TAB"
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opPushLiteral, _
                                              " ", _
                                              "Tab function: push space for the string op") Then Return compiler__parseFail_("functionCall")
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opPushLiteral, _
                                              vbTab, _
                                              "Make the tab") Then Return compiler__parseFail_("functionCall")
                Case "TRIM"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opTrim, _
                                              0, _
                                              "Trim string (remove leading and trailing spaces") Then Return compiler__parseFail_("functionCall")
                Case "UBOUND"
                    If Not compiler__functionCall__bounds_(intIndex, _
                                                           objScanned, _
                                                           colPolish, _
                                                           intArgEnd, _
                                                           strSourceCode, _
                                                           True, _
                                                           colVariables, _
                                                           intLevel + 1) Then Return compiler__parseFail_("functionCall")
                Case "UCASE"
                    If Not compiler__expressionList_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intArgEnd, _
                                                     strSourceCode, _
                                                     1, _
                                                     1, _
                                                     colVariables, _
                                                     intParms, _
                                                     intLevel + 1) Then
                        Return compiler__parseFail_("functionCall")
                    End If
                    If Not compiler__genCode_(colPolish, intIndex, ENUop.opUCase, 0, "Replace string by upper case") Then Return compiler__parseFail_("functionCall")
                Case Else
                    intIndex = intIndex1
                    Return compiler__parseFail_("functionCall")
            End Select
            intIndex = intArgEnd + 1
            booSideEffects = True
            compiler__parseEvent_("functionCall", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        compiler__parseFail_("functionCall")
    End Function

    ' ----------------------------------------------------------------------
    ' Compile LBound and UBound on behalf of compiler__functionCall_
    '
    '
    Private Function compiler__functionCall__bounds_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByVal booUBound As Boolean, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        Dim intSymIndex As Integer
        If Not compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intSymIndex, _
                                 colVariables, _
                                 False, _
                                 intLevel + 1) Then
            compiler__errorHandler_("Invalid LBound/UBound statement does not validly identify an array", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode, _
                                    "")
            Return compiler__parseFail_("functionCallBounds")
        End If
        Dim intDimension As Integer = 1
        Dim objDimension As qbVariable.qbVariable
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 ",", _
                                 intLevel + 1) Then
            If Not compiler__expression_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         objDimension, _
                                         colVariables, _
                                         intLevel + 1) Then
                compiler__errorHandler_("Invalid LBound/UBound statement fails to properly specify the dimension", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "The dimension must indicate the bound to be returned from 1")
                Return compiler__parseFail_("functionCallBounds")
            End If
            Try
                intDimension = CInt(objDimension.value)
            Catch
                compiler__errorHandler_("Invalid LBound/UBound statement specifies the non-integer dimension " & _
                                        _OBJutilities.object2String(objDimension), _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "The dimension must indicate the bound to be returned from 1")
                Return compiler__parseFail_("functionCallBounds")
            End Try
        End If
        If intIndex <> intEndIndex Then
            compiler__errorHandler_("Invalid LBound/UBound statement contains unrecognizable argument code", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode, _
                                    "")
            Return compiler__parseFail_("functionCallBounds")
        End If
        Dim intBound As Integer
        With CType(colVariables.Item(intSymIndex), qbVariable.qbVariable)
            If .Dope.Dimensions < 1 Then
                compiler__errorHandler_("Variable has no dimensions defined.  LBound and UBound can only be used with " & _
                                        "variables explicitly defined using Dim statements.", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "")
                Return compiler__parseFail_("functionCallBounds")
            End If
            If intDimension < 0 OrElse intDimension > .Dope.Dimensions Then
                compiler__errorHandler_("The array " & .VariableName & " does not have the dimension " & intDimension, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        "")
                Return compiler__parseFail_("functionCallBounds")
            End If
            If booUBound Then
                intBound = .Dope.UpperBound(intDimension)
            Else
                ' At this time, the lowerBound is always zero
                intBound = 0
            End If
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, intBound, _
                                        "Push the known " & _
                                        CStr(IIf(booUBound, "upper", "lower")) & " " & _
                                        "bound of the array " & _
                                        _OBJutilities.enquote(.VariableName) & _
                                        CStr(IIf(.Dope.Dimensions = 1, _
                                                 "", _
                                                 "(at dimension " & intDimension & ")"))) Then
                Return compiler__parseFail_("functionCallBounds")
            End If
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Generate code to the Polish notation output
    '
    '
    ' --- No source length or comment: length defaults to 1: comment to op desc
    Private Overloads Function compiler__genCode_(ByRef colPolish As Collection, _
                                                  ByVal intIndex As Integer, _
                                                  ByVal enuOpCode As qbOp.qbOp.ENUop, _
                                                  ByVal objOperand As Object) As Boolean
        Return (compiler__genCode_(colPolish, enuOpCode, intIndex, 1, objOperand, ""))
    End Function
    ' --- No source length or comment: length defaults to 1: comment to op desc
    Private Overloads Function compiler__genCode_(ByRef colPolish As Collection, _
                                                  ByVal intIndex As Integer, _
                                                  ByVal enuOpCode As qbOp.qbOp.ENUop, _
                                                  ByVal objOperand As Object, _
                                                  ByVal strComment As String) As Boolean
        Return (compiler__genCode_(colPolish, enuOpCode, intIndex, 1, objOperand, strComment))
    End Function
    ' --- Length, no comment
    Private Overloads Function compiler__genCode_(ByRef colPolish As Collection, _
                                                  ByVal enuOpCode As qbOp.qbOp.ENUop, _
                                                  ByVal intIndex As Integer, _
                                                  ByVal intLength As Integer, _
                                                  ByVal objOperand As Object) As Boolean
        Return (compiler__genCode_(colPolish, enuOpCode, intIndex, 1, objOperand, ""))
    End Function
    ' --- Complete info
    Private Overloads Function compiler__genCode_(ByRef colPolish As Collection, _
                                                  ByVal enuOpCode As ENUop, _
                                                  ByVal intIndex As Integer, _
                                                  ByVal intLength As Integer, _
                                                  ByVal objOperand As Object, _
                                                  ByVal strComment As String) As Boolean
        Try
#If QUICKBASICENGINE_EXTENSION Then
            If Not OBJstate.usrState.booGenerateNOPs _
               AndAlso _
               (enuOpCode = ENUop.opNop OrElse enuOpCode = ENUop.opRem) Then
                Return True
            End If
#End If
            Dim objNew As qbPolish.qbPolish
            objNew = New qbPolish.qbPolish
            With objNew
                .Opcode = enuOpCode
                .Operand = objOperand
                .Comment = CStr(IIf(strComment = "", _
                                    _OBJqbOp.opCodeToDescription(enuOpCode), _
                                    strComment))
                .TokenStartIndex = intIndex
                .TokenLength = intLength
            End With
            colPolish.Add(objNew)
            raiseEvent_("codeGenEvent", objNew)
        Catch objException As Exception
            errorHandler_("Internal compiler error: cannot expand the output code: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__genCode_", "Marking object unusable: returning False", _
                          objException)
            Return (False)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Generate code for assignment (implicit, explicit, Input statement, 
    ' and read
    '
    '
    Private Function compiler__generateAssignment_(ByVal intIndex As Integer, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intMemoryLocation As Integer, _
                                                   ByVal strComment As String) As Boolean
        Dim enuOperator As ENUop
        Dim strCommentWork As String = strComment
        Dim strStackInfo As String
        Select Case intMemoryLocation
            Case MEMORY_LOCATION_STACKEDADDRESS
                strStackInfo = "address for indirect assignment expected"
                enuOperator = ENUop.opPopIndirect
            Case MEMORY_LOCATION_STACKFRAME
                strStackInfo = "array stack frame and value are expected"
                enuOperator = ENUop.opPopToArrayElement
            Case Else
                enuOperator = ENUop.opPop
        End Select
        _OBJutilities.append(strCommentWork, ": ", strStackInfo)
        If compiler__genCode_(colPolish, intIndex, _
                                enuOperator, _
                                intMemoryLocation, _
                                strCommentWork) Then
            If intMemoryLocation = -1 Then
                If Not compiler__genCode_(colPolish, intIndex, _
                                          ENUop.opPop, _
                                          0, _
                                          "Remove indirect address") Then
                    compiler__parseFail_("implicitAssignment")
                End If
            End If
            Return (True)
        End If
    End Function
    ' ----------------------------------------------------------------------
    ' Get one label, increase sequence number
    '
    '
    Private Function compiler__getLabel_() As String
        With OBJstate.usrState
            .intLabelSeq += 1
            Return ("LBL" & .intLabelSeq)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Get one token's value
    '
    '
    Private Function compiler__getToken_(ByVal intIndex As Integer, _
                                         ByVal objScanned As qbScanner.qbScanner, _
                                         ByVal strSourceCode As String) As String
        With objScanned.QBToken(intIndex)
            Return Mid(strSourceCode, .StartIndex, .Length)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' immediateCommand := singleImmediateCommand (:singleImmediateCommand)*
    '
    '
    Private Function compiler__immediateCommand_(ByRef intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByRef colPolish As Collection, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByVal strSourceCode As String, _
                                                 ByRef colVariables As Collection, _
                                                 ByRef objImmediateValue As qbVariable.qbVariable, _
                                                 ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "immediateCommand")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex2 As Integer = intIndex
        If Not compiler__singleImmediateCommand_(intIndex, _
                                                 objScanned, _
                                                 colPolish, _
                                                 intEndIndex, _
                                                 strSourceCode, _
                                                 colVariables, _
                                                 objImmediateValue, _
                                                 intLevel + 1) Then
            Return compiler__parseFail_("immediateCommand")
        End If
        Dim intIndex1 As Integer
        Do
            intIndex1 = intIndex
            If Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon, _
                                         intLevel + 1) Then
                Exit Do
            End If
            If Not compiler__singleImmediateCommand_(intIndex, _
                                                     objScanned, _
                                                     colPolish, _
                                                     intEndIndex, _
                                                     strSourceCode, _
                                                     colVariables, _
                                                     objImmediateValue, _
                                                     intLevel + 1) Then
                intIndex = intIndex1 : Exit Do
            End If
        Loop
        compiler__parseEvent_("immediateCommand", _
                                False, _
                                intIndex2, _
                                intIndex - intIndex2, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' implicitAssignment := lValue = expression
    '
    '
    Private Function compiler__implicitAssignment_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef colVariables As Collection, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "implicitAssignment")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim intIndex2 As Integer = intIndex
        Dim intMemoryLocation As Integer
        If Not compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intMemoryLocation, _
                                 colVariables, _
                                 False, _
                                 intLevel + 1) Then
            Return compiler__parseFail_("implicitAssignment")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "=", _
                                     intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("implicitAssignment")
        End If
        Dim strLValue As String = rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 2)
        intIndex1 = intIndex
        Dim objConstantValue As qbVariable.qbVariable
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValue, _
                                     colVariables, _
                                     intLevel + 1) Then
            intIndex = intIndex1
            compiler__parseFail_("implicitAssignment")
        End If
        Dim strExpression As String = rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1)
        If Not (objConstantValue Is Nothing) Then
            If Not compiler__genCode_(colPolish, intIndex, _
                                      ENUop.opPushLiteral, _
                                      objConstantValue, _
                                      "Assign value " & _OBJutilities.object2String(objConstantValue) & " " & _
                                      "of " & _OBJutilities.enquote(strExpression) & " " & _
                                      "to " & _OBJutilities.enquote(strLValue)) Then Return compiler__parseFail_("implicitAssignment")
        End If
        If Not compiler__generateAssignment_(intIndex, _
                                             colPolish, _
                                             intMemoryLocation, _
                                             "Assign value of " & _
                                             _OBJutilities.enquote(strExpression) & _
                                             "to " & _
                                             strLValue) Then Return False
        compiler__parseEvent_("implicitAssignment", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Return True (word to yo Mama is reserved) or False
    '
    '
    Private Function compiler__isReservedWord_(ByVal strWord As String) As Boolean
        Return InStr(" " & QB_RESERVED_WORDS & " ", " " & LCase(strWord) & " ") <> 0
    End Function

    ' ----------------------------------------------------------------------
    ' likeFactor := concatFactor [likeFactorRHS]
    '
    '
    Private Function compiler__likeFactor_(ByRef intIndex As Integer, _
                                           ByVal objScanned As qbScanner.qbScanner, _
                                           ByRef colPolish As Collection, _
                                           ByVal intEndIndex As Integer, _
                                           ByVal strSourceCode As String, _
                                           ByRef objConstantValue As qbVariable.qbVariable, _
                                           ByRef colVariables As Collection, _
                                           ByRef booSideEffects As Boolean, _
                                           ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "likeFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__concatFactor_(intIndex, _
                                       objScanned, _
                                       colPolish, _
                                       intEndIndex, _
                                       strSourceCode, _
                                       objConstantValue, _
                                       colVariables, _
                                       booSideEffects, _
                                       intLevel + 1) Then
            Return compiler__parseFail_("likeFactor")
        End If
        If intIndex <= intEndIndex Then
            Dim intIndex2 As Integer = intIndex
            If Not compiler__likeFactorRHS_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            booSideEffects, _
                                            intLevel + 1) _
                AndAlso _
                intIndex2 <> intIndex Then Return compiler__parseFail_("likeFactor")
        End If
        compiler__parseEvent_("likeFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' likeFactorRHS := & concatFactor [likefactorRHS]
    '
    '
    Private Function compiler__likeFactorRHS_(ByRef intIndex As Integer, _
                                              ByVal objScanned As qbScanner.qbScanner, _
                                              ByRef colPolish As Collection, _
                                              ByVal intEndIndex As Integer, _
                                              ByVal strSourceCode As String, _
                                              ByRef objConstantValue As qbVariable.qbVariable, _
                                              ByRef colVariables As Collection, _
                                              ByRef booSideEffects As Boolean, _
                                              ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "likeFactorRHS")
        If intIndex > intEndIndex Then Return compiler__parseFail_("likeFactorRHS")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim strRelOp As String
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand, _
                                     intLevel + 1) Then Return compiler__parseFail_("likeFactorRHS")
        Dim objConstantValueRHS As qbVariable.qbVariable
        Dim booLHSsideEffects As Boolean = booSideEffects
        If Not compiler__concatFactor_(intIndex, _
                                       objScanned, _
                                       colPolish, _
                                       intEndIndex, _
                                       strSourceCode, _
                                       objConstantValueRHS, _
                                       colVariables, _
                                       booSideEffects, _
                                       intLevel + 1) Then
            Return compiler__parseFail_("likeFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_("&", _
                                                  booLHSsideEffects, _
                                                  objConstantValue, _
                                                  booSideEffects, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__likeFactorRHS_(intIndex, _
                                        objScanned, _
                                        colPolish, _
                                        intEndIndex, _
                                        strSourceCode, _
                                        objConstantValue, _
                                        colVariables, _
                                        booSideEffects, _
                                        intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("likeFactorRHS")
        compiler__parseEvent_("likeFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Check for a newline including a colon newline
    '
    '
    Private Function compiler__logicalNewline_(ByRef intIndex As Integer, _
                                               ByVal objScanned As qbScanner.qbScanner, _
                                               ByVal strSourceCode As String, _
                                               ByVal intEndIndex As Integer, _
                                               ByVal intLevel As Integer) As Boolean
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline, _
                                 intLevel + 1) Then Return (True)
        Return compiler__checkToken_(intIndex, _
                                      objScanned, _
                                      strSourceCode, _
                                      intEndIndex, _
                                      qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon, _
                                      intLevel + 1) 
    End Function

    ' ----------------------------------------------------------------------
    ' lValue := typedIdentifier [ parenthesizedSubscript ]
    '
    '
    ' Note that this routine returns one of these values in intMemoryReference:
    '
    '
    '      *  A memory reference
    '
    '      *  -1 indicates that at run time the address will be
    '         dynamically calculated and available at the top of
    '         the stack
    '
    '
    ' In the overload that includes a ByRef symIndex the symIndex is set
    ' to the index of the lValue in the symbol table, unless the memory
    ' reference is dynamic, where symIndex is not changed.
    '
    '
    ' --- Drop symbol's index
    Private Overloads Function compiler__lValue_(ByRef intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByRef colPolish As Collection, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByVal strSourceCode As String, _
                                                 ByRef intMemoryReference As Integer, _
                                                 ByRef colVariables As Collection, _
                                                 ByVal booDefContext As Boolean, _
                                                 ByVal intLevel As Integer) As Boolean
        Dim intSymIndex As Integer
        Return compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intMemoryReference, _
                                 intSymIndex, _
                                 colVariables, _
                                 booDefContext, _
                                 intLevel)
    End Function
    ' --- Assign symbol's index
    Private Overloads Function compiler__lValue_(ByRef intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByRef colPolish As Collection, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByVal strSourceCode As String, _
                                                 ByRef intMemoryReference As Integer, _
                                                 ByRef intSymIndex As Integer, _
                                                 ByRef colVariables As Collection, _
                                                 ByVal booDefContext As Boolean, _
                                                 ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "lValue")
        Dim objType As qbVariableType.qbVariableType
        Try
            objType = New qbVariableType.qbVariableType
        Catch objException As Exception
            errorHandler_("Internal compiler error: cannot create variableType: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__lValue_", _
                          "Marking object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False
            Return compiler__parseFail_("lValue")
        End Try
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim strIdentifier As String
        If Not compiler__typedIdentifier_(intIndex, _
                                          objScanned, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          strIdentifier, _
                                          objType, _
                                          intLevel + 1) Then
            Return compiler__parseFail_("lValue")
        End If
        If compiler__isReservedWord_(strIdentifier) Then
            intIndex = intIndex1
            Return compiler__parseFail_("lValue")
        End If
        If Not compiler__findOrAddVariable_(strIdentifier, _
                                            intSymIndex, _
                                            colVariables, _
                                            objScanned, _
                                            strSourceCode, _
                                            booDefContext) Then Return compiler__parseFail_("lValue")
        intMemoryReference = intSymIndex   ' Always the same for this version of the compiler
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "(", _
                                 intLevel + 1) Then
            Dim intRPindex As Integer = _
                objScanned.findRightParenthesis(intIndex, _
                                                intEndIndex:=intEndIndex)
            If Not compiler__expressionList_(intIndex, _
                                             objScanned, _
                                             colPolish, _
                                             intRPindex, _
                                             strSourceCode, _
                                             intMemoryReference, _
                                             colVariables, _
                                             intSymIndex, _
                                             intLevel + 1) Then Return compiler__parseFail_("lValue")
            intIndex = intRPindex + 1
            intMemoryReference = MEMORY_LOCATION_STACKFRAME
        End If
        If booDefContext Then
            compiler__asClause_(intIndex, _
                                objScanned, _
                                intEndIndex, _
                                strSourceCode, _
                                intMemoryReference, _
                                colVariables, _
                                intLevel)
        End If
        compiler__parseEvent_("lValue", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' lValueList = lValue [ comma lValueList ]
    '
    '
    ' The enuContext parameter specifies the context of the lValue list:
    '
    '
    '      *  definition: lValue list appears as part of a Dim statement
    '      *  input: lValue list appears as part of any Input statement
    '      *  read: lValue list appears as part of any Read statement
    '
    '
    Private Enum ENUlValueContext
        definition
        input
        read
    End Enum
    Private Function compiler__lValueList_(ByRef intIndex As Integer, _
                                           ByVal objScanned As qbScanner.qbScanner, _
                                           ByRef colPolish As Collection, _
                                           ByVal intEndIndex As Integer, _
                                           ByVal strSourceCode As String, _
                                           ByVal enuContext As ENUlValueContext, _
                                           ByRef colVariables As Collection, _
                                           ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "lValueList")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Select Case enuContext
            Case ENUlValueContext.definition
                If Not compiler__lValueNew_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            colVariables, _
                                            intLevel + 1) Then
                    Return compiler__parseFail_("lValueList")
                End If
            Case ENUlValueContext.input
                If Not compiler__lValueList__inputRead_(intIndex, _
                                                        objScanned, _
                                                        colPolish, _
                                                        intEndIndex, _
                                                        strSourceCode, _
                                                        colVariables, _
                                                        True, _
                                                        intLevel + 1) Then Return compiler__parseFail_("lValueList")
            Case ENUlValueContext.read
                If Not compiler__lValueList__inputRead_(intIndex, _
                                                        objScanned, _
                                                        colPolish, _
                                                        intEndIndex, _
                                                        strSourceCode, _
                                                        colVariables, _
                                                        False, _
                                                        intLevel + 1) Then Return compiler__parseFail_("lValueList")
            Case Else
                errorHandler_("Internal compiler error: unexpected lValue context", _
                              "compiler__lValueList_", _
                              "Marking object unusable and returning False", _
                              Nothing)
                Return compiler__parseFail_("lValueList")
        End Select
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma, _
                                     intLevel + 1) _
           OrElse _
           compiler__lValueList_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 enuContext, _
                                 colVariables, _
                                 intLevel + 1) Then
            compiler__parseEvent_("lValueList", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("lValueList")
    End Function

    ' ----------------------------------------------------------------------
    ' Compile the input and the read statements
    '
    '
    Private Function compiler__lValueList__inputRead_(ByRef intIndex As Integer, _
                                                      ByVal objScanned As qbScanner.qbScanner, _
                                                      ByRef colPolish As Collection, _
                                                      ByVal intEndIndex As Integer, _
                                                      ByVal strSourceCode As String, _
                                                      ByRef colVariables As Collection, _
                                                      ByVal booInput As Boolean, _
                                                      ByVal intLevel As Integer) As Boolean
        Dim intMemoryReference As Integer
        Dim intStartIndex As Integer = intIndex
        If Not compiler__lValue_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    intMemoryReference, _
                                    colVariables, _
                                    False, _
                                    intLevel) Then
            Return (False)
        End If
        Dim strSource As String = rebuildCode_(objScanned, strSourceCode, intStartIndex, intIndex - 1)
        If Not compiler__genCode_(colPolish, _
                                  intIndex, _
                                  CType(IIf(booInput, _
                                            ENUop.opInput, _
                                            ENUop.opRead), _
                                        qbOp.qbOp.ENUop), _
                                  Nothing, _
                                  "Read from " & _
                                  CStr(IIf(booInput, "standard input", "data")) & " " & _
                                       "to stack(top)") Then Return (False)
        Return (compiler__generateAssignment_(intIndex, colPolish, intMemoryReference, _
                                              "Assign input or read statement value to storage"))
    End Function

    ' ----------------------------------------------------------------------
    ' Parse lValue (where it MUST be a new location in a Dim)
    '
    '
    Private Function compiler__lValueNew_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef colVariables As Collection, _
                                          ByVal intLevel As Integer) As Boolean
        Dim intMemoryReference As Integer
        Dim intOldStorageSize As Integer = colVariables.Count
        If Not compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intMemoryReference, _
                                 colVariables, _
                                 True, _
                                 intLevel) Then
            Return (False)
        End If
        Return (colVariables.Count > intOldStorageSize)
    End Function

    ' ----------------------------------------------------------------------
    ' moduleDefinition := subDefinition | functionDefinition
    '
    '
    Private Function compiler__moduleDefinition_(ByRef intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByRef colPolish As Collection, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByVal strSourceCode As String, _
                                                 ByRef colVariables As Collection, _
                                                 ByRef objNestingStack As Stack, _
                                                 ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "moduleDefinition")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__moduleDefinition__("Sub", _
                                           intIndex, _
                                           objScanned, _
                                           colPolish, _
                                           intEndIndex, _
                                           strSourceCode, _
                                           colVariables, _
                                           objNestingStack, _
                                           intLevel + 1) _
              OrElse _
              compiler__moduleDefinition__("Function", _
                                           intIndex, _
                                           objScanned, _
                                           colPolish, _
                                           intEndIndex, _
                                           strSourceCode, _
                                           colVariables, _
                                           objNestingStack, _
                                           intLevel + 1) Then
            compiler__parseEvent_("moduleDefinition", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
       End If
       compiler__parseFail_("moduleDefinition")
    End Function

    ' ----------------------------------------------------------------------
    ' subHeader := Sub identifier [ formalParameterList ]
    ' functionHeader := Function identifier [ formalParameterList ]
    '
    '
    Private Function compiler__moduleDefinition__(ByVal strSubFunction As String, _
                                                  ByRef intIndex As Integer, _
                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                  ByRef colPolish As Collection, _
                                                  ByVal intEndIndex As Integer, _
                                                  ByVal strSourceCode As String, _
                                                  ByRef colVariables As Collection, _
                                                  ByRef objNestingStack As Stack, _
                                                  ByVal intLevel As Integer) As Boolean
        Dim strGC As String
        Select Case UCase(Trim(strSubFunction))
            Case "SUB" : strGC = "subHeader"
            Case "FUNCTION" : strGC = "functionHeader"
            Case Else
                errorHandler_("Internal compiler error: strSubFunction is not valid", _
                              "compiler__moduleDefinition__", _
                              "Marking object not usable and returning False", _
                              Nothing)
                OBJstate.usrState.booUsable = False
                Return False
        End Select
        raiseEvent_("parseStartEvent", strGC)
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     strSubFunction, _
                                     intLevel + 1) Then Return compiler__parseFail_(strGC)
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     objScanned.QBToken(intIndex).TokenType, _
                                     intLevel + 1) Then
            intIndex = intIndex1
            Return compiler__parseFail_(strGC)
        End If
        Try
            ReDim Preserve OBJstate.usrState.usrSubFunction(UBound(OBJstate.usrState.usrSubFunction) + 1)
        Catch
            errorHandler_("Cannot expand the Sub/Function table: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__moduleDefinition__", _
                          "Making object unusable and returning False", _
                          Nothing)
            OBJstate.usrState.booUsable = False
            Return compiler__parseFail_(strGC)
        End Try
        With OBJstate.usrState.usrSubFunction(UBound(OBJstate.usrState.usrSubFunction))
            .booFunction = (UCase(strSubFunction) = "FUNCTION")
            .intLocation = colPolish.Count + 1
            If Not compiler__formalParameterList_(intIndex, _
                                                  objScanned, _
                                                  colPolish, _
                                                  intEndIndex, _
                                                  strSourceCode, _
                                                  colVariables, _
                                                  objNestingStack, _
                                                  .usrFormalParameters, _
                                                  intLevel + 1) Then
                intIndex = intIndex1
                Return compiler__parseFail_(strGC)
            End If
            Dim colIndexEntry As Collection
            Try
                colIndexEntry = New Collection
                With colIndexEntry
                    .Add(strSubFunction) : .Add(UBound(OBJstate.usrState.usrSubFunction))
                End With
            Catch objException As Exception
                errorHandler_("Cannot create index collection entry: " & _
                              Err.Number & " " & Err.Description, _
                              "compiler__moduleDefinition__", _
                              "Making object unusable and returning False", _
                              objException)
                OBJstate.usrState.booUsable = False : Return compiler__parseFail_(strGC)
            End Try
            compiler__parseEvent_(strGC, _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        Return compiler__parseFail_(strGC)
    End Function

    ' ----------------------------------------------------------------------
    ' mulFactor := powFactor [mulFactorRHS]
    '
    '
    Private Function compiler__mulFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "mulFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__powFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("mulFactor")
        End If
        If intIndex <= intEndIndex Then
            Dim intIndex2 As Integer = intIndex
            If Not compiler__mulFactorRHS_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            booSideEffects, _
                                            intLevel + 1) _
                AndAlso _
                intIndex2 <> intIndex Then Return compiler__parseFail_("mulFactor")
        End If
        compiler__parseEvent_("mulFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' mulFactorRHS = powOp powFactor [compiler__mulFactorRHS_]
    '
    '
    Private Function compiler__mulFactorRHS_(ByRef intIndex As Integer, _
                                             ByVal objScanned As qbScanner.qbScanner, _
                                             ByRef colPolish As Collection, _
                                             ByVal intEndIndex As Integer, _
                                             ByVal strSourceCode As String, _
                                             ByRef objConstantValue As qbVariable.qbVariable, _
                                             ByRef colVariables As Collection, _
                                             ByRef booSideEffects As Boolean, _
                                             ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "mulFactorRHS")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If intIndex > intEndIndex Then Return compiler__parseFail_("mulFactorRHS")
        Dim strPowOp As String
        If Not compiler__powOp_(objScanned, intIndex, strSourceCode, strPowOp, intLevel) Then Return compiler__parseFail_("mulFactorRHS")
        Dim booSideEffectsRHS As Boolean
        Dim objConstantValueRHS As qbVariable.qbVariable
        If Not compiler__powFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValueRHS, _
                                    colVariables, _
                                    booSideEffectsRHS, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("mulFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_(strPowOp, _
                                                  booSideEffects, _
                                                  objConstantValue, _
                                                  booSideEffectsRHS, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__mulFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("mulFactorRHS")
        compiler__parseEvent_("mulFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' A multiply operator is multiply, divide or integer divide (they have 
    ' the same precedence)
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression.
    '
    '
    Private Function compiler__mulOp_(ByVal objScanned As qbScanner.qbScanner, _
                                      ByRef intIndex As Integer, _
                                      ByVal strSourceCode As String, _
                                      ByRef strOp As String, _
                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "mulOp")
        strOp = UCase(Mid(strSourceCode, _
                          objScanned.QBToken(intIndex).StartIndex, _
                          objScanned.QBToken(intIndex).Length))
        Select Case strOp
            Case "*"
            Case "/"
            Case "\"
            Case "MOD"
            Case Else : Return compiler__parseFail_("mulOp")
        End Select
        intIndex += 1       ' I do miss C at times!
        compiler__parseEvent_("mulOp", _
                                False, _
                                intIndex - 1, _
                                1, _
                                0, 0, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Nesting level to string
    '
    '
    Private Function compiler__nesting2String_(ByVal usrNesting As TYPnesting) As String
        With usrNesting
            Return (compiler__nestingType2String_(.enuNestType) & " " & _
                   "at line " & .intLine)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Convert the nesting type enum to a string
    '
    '
    Private Function compiler__nestingType2String_(ByVal enuNest As ENUnesting) As String
        Select Case enuNest
            Case ENUnesting.doLoop : Return ("Do")
            Case ENUnesting.elseClause : Return ("Else")
            Case ENUnesting.forLoop : Return ("For")
            Case ENUnesting.ifThen : Return ("If")
            Case ENUnesting.openCode : Return ("Open code")
            Case ENUnesting.whileLoop : Return ("While")
            Case ENUnesting.elseClause : Return ("?")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' notFactor := likeFactor [notFactorRHS]
    '
    '
    Private Function compiler__notFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "notFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__likeFactor_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValue, _
                                     colVariables, _
                                     booSideEffects, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("notFactor")
        End If
        If intIndex <= intEndIndex Then
            Dim intIndex2 As Integer = intIndex
            If Not compiler__notFactorRHS_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            booSideEffects, _
                                            intLevel + 1) _
                AndAlso _
                intIndex2 <> intIndex Then Return compiler__parseFail_("notFactor")
        End If
        compiler__parseEvent_("notFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' notFactorRHS := Like likeFactor [notFactorRHS]
    '
    '
    Private Function compiler__notFactorRHS_(ByRef intIndex As Integer, _
                                                ByVal objScanned As qbScanner.qbScanner, _
                                                ByRef colPolish As Collection, _
                                                ByVal intEndIndex As Integer, _
                                                ByVal strSourceCode As String, _
                                                ByRef objConstantValue As qbVariable.qbVariable, _
                                                ByRef colVariables As Collection, _
                                                ByRef booSideEffects As Boolean, _
                                                ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "notFactorRHS")
        Dim intCountPrevious As Integer = colPolish.Count
        If intIndex > intEndIndex Then Return compiler__parseFail_("notFactorRHS")
        Dim intIndex1 As Integer = intIndex
        Dim strRelOp As String
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "Like", _
                                     intLevel + 1) Then Return compiler__parseFail_("notFactorRHS")
        Dim objConstantValueRHS As qbVariable.qbVariable
        Dim booLHSsideEffects As Boolean = booSideEffects
        If Not compiler__likeFactor_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantValueRHS, _
                                     colVariables, _
                                     booSideEffects, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("notFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_("Like", _
                                                  booLHSsideEffects, _
                                                  objConstantValue, _
                                                  booSideEffects, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__notFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then compiler__parseFail_("notFactorRHS")
        compiler__parseEvent_("notFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' number := (+|-){0-1} ( realNumber | integer ) [ numTypeChar ]
    '
    '
    Private Function compiler__number_(ByRef intIndex As Integer, _
                                       ByVal objScanned As qbScanner.qbScanner, _
                                       ByVal intEndIndex As Integer, _
                                       ByVal strSourceCode As String, _
                                       ByRef dblValue As Double, _
                                       ByRef objType As qbVariableType.qbVariableType, _
                                       ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "number")
        Dim intIndex1 As Integer = intIndex
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger, _
                                 intLevel + 1) _
           OrElse _
           compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber, _
                                 intLevel + 1) Then
            dblValue = CDbl(Mid(strSourceCode, _
                                objScanned.QBToken(intIndex - 1).StartIndex, _
                                objScanned.QBToken(intIndex - 1).Length))
            compiler__numTypeChar_(intIndex, _
                                   objScanned, _
                                   intEndIndex, _
                                   strSourceCode, _
                                   objType, _
                                   intLevel + 1)
            compiler__parseEvent_("number", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    0, 0, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("number")
    End Function

    ' ----------------------------------------------------------------------
    ' numTypeChar = percent | ampersand | exclamation | poundSign
    '
    '
    ' Note that the numeric type character must be adjacent to the number.
    Private Function compiler__numTypeChar_(ByRef intIndex As Integer, _
                                            ByVal objScanned As qbScanner.qbScanner, _
                                            ByVal intEndIndex As Integer, _
                                            ByVal strSourceCode As String, _
                                            ByRef objType As qbVariableType.qbVariableType, _
                                            ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "numTypeChar")
        Dim intIndex1 As Integer = intIndex
        Try
            objType = New qbVariableType.qbVariableType("Variant")
        Catch objException As Exception
            errorHandler_("Can't make qbVariableType: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__numTypeChar_", _
                          "Returning False", _
                          objException)
            Return compiler__parseFail_("numTypeChar")
        End Try
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand, _
                                 intLevel + 1) _
           OrElse _
           compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeExclamation, _
                                 intLevel + 1) _
           OrElse _
           compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypePercent, _
                                 intLevel + 1) _
           OrElse _
           compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypePound, _
                                 intLevel + 1) Then
            If objScanned.tokenEndIndex(intIndex - 2) + 1 = objScanned.tokenStartIndex(intIndex - 1) Then
                Dim intIndex2 As Integer = intIndex - 1
                With objScanned.QBToken(intIndex2)
                    Dim strType As String = compiler__typechar2Type_(intIndex, _
                                                                     objScanned, _
                                                                     strSourceCode, _
                                                                     Mid(strSourceCode, _
                                                                         .StartIndex, _
                                                                         .Length))
                    objType.fromString(strType)
                End With
                compiler__parseEvent_("numTypeChar", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        0, 0, _
                                        "", _
                                        intLevel)
                Return (True)
            End If
        End If
        intIndex = intIndex1
        Return compiler__parseFail_("numTypeChar")
    End Function

    ' ----------------------------------------------------------------------
    ' openCode := statement [ logicalNewline [ sourceProgramBody ] ]
    '
    '
    Private Function compiler__openCode_(ByRef intIndex As Integer, _
                                         ByVal objScanned As qbScanner.qbScanner, _
                                         ByRef colPolish As Collection, _
                                         ByVal intEndIndex As Integer, _
                                         ByVal strSourceCode As String, _
                                         ByRef colVariables As Collection, _
                                         ByRef objNestingStack As Stack, _
                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "openCode")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__statement_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objNestingStack, _
                                    colVariables, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("openCode")
        End If
        If compiler__logicalNewline_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         intLevel) _
           AndAlso _
           intIndex <= intEndIndex _
           AndAlso _
           Not compiler__sourceProgramBody_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            colVariables, _
                                            objNestingStack, _
                                            intLevel + 1) Then
            intIndex = intIndex1
            Return compiler__parseFail_("openCode")
        End If
        compiler__parseEvent_("openCode", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' optionStmt := Option Explicit 
    '
    '
    Private Function compiler__optionStmt_(ByRef intIndex As Integer, _
                                           ByVal objScanned As qbScanner.qbScanner, _
                                           ByRef colPolish As Collection, _
                                           ByVal intEndIndex As Integer, _
                                           ByVal strSourceCode As String, _
                                           ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "optionStmt")
        If intIndex > intEndIndex Then compiler__parseFail_("optionStmt")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "OPTION", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("optionStmt")
        End If
        With OBJstate.usrState
            If compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "EXPLICIT", _
                                     intLevel + 1) Then
                .booExplicit = True
                compiler__parseEvent_("optionStmt", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "", _
                                        intLevel)
                Return (True)
            End If
            compiler__errorHandler_("Invalid Option statement", intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("optionStmt")
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' orFactor := andFactor [ orFactorRHS ]
    '
    '
    Private Function compiler__orFactor_(ByRef intIndex As Integer, _
                                         ByVal objScanned As qbScanner.qbScanner, _
                                         ByRef colPolish As Collection, _
                                         ByVal intEndIndex As Integer, _
                                         ByVal strSourceCode As String, _
                                         ByRef objConstantValue As qbVariable.qbVariable, _
                                         ByRef colVariables As Collection, _
                                         ByRef booSideEffects As Boolean, _
                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "orFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__andFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("orFactor")
        End If
        If intIndex <= intEndIndex Then
            Dim intIndex2 As Integer = intIndex
            If Not compiler__orFactorRHS_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            booSideEffects, _
                                            intLevel + 1) _
               AndAlso _
               intIndex2 <> intIndex Then compiler__parseFail_("orFactor")
        End If
        compiler__parseEvent_("orFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' orFactorRHS = andOp andFactor [orFactorRHS]
    '
    '
    Private Function compiler__orFactorRHS_(ByRef intIndex As Integer, _
                                                ByVal objScanned As qbScanner.qbScanner, _
                                                ByRef colPolish As Collection, _
                                                ByVal intEndIndex As Integer, _
                                                ByVal strSourceCode As String, _
                                                ByRef objConstantValue As qbVariable.qbVariable, _
                                                ByRef colVariables As Collection, _
                                                ByRef booSideEffects As Boolean, _
                                                ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "orFactorRHS")
        If intIndex > intEndIndex Then
            Return compiler__parseFail_("orFactorRHS")
        End If
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim strOp As String = compiler__getToken_(intIndex, objScanned, strSourceCode)
        If Not compiler__andOp_(objScanned, _
                                intIndex, _
                                strSourceCode, _
                                strOp, _
                                intLevel + 1) Then Return compiler__parseFail_("orFactorRHS")
        Dim strLbl As String
        If UCase(strOp) = "ANDALSO" Then
            strLbl = compiler__getLabel_()
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opDuplicate, _
                                      Nothing, _
                                      "AndAlso: duplicate stack and skip RHS when LHS is False") _
               OrElse _
               Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opJumpZ, _
                                      strLbl, _
                                      "") Then
                Return compiler__parseFail_("orFactorRHS")
            End If
        End If
        Dim booSideEffectsRHS As Boolean
        Dim objConstantValueRHS As qbVariable.qbVariable
        If Not compiler__andFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValueRHS, _
                                    colVariables, _
                                    booSideEffectsRHS, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("orFactorRHS")
        End If
        If UCase(strOp) = "ANDALSO" Then
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLbl, "AndAlso jump target for False") Then
                Return compiler__parseFail_("orFactorRHS")
            End If
        Else
            objConstantValue = compiler__binaryOpGen_("AND", _
                                                      booSideEffects, _
                                                      objConstantValue, _
                                                      booSideEffectsRHS, _
                                                      objConstantValueRHS, _
                                                      colPolish, _
                                                      intIndex)
        End If
        Dim intIndex2 As Integer = intIndex
        If Not compiler__orFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("orFactorRHS")
        compiler__parseEvent_("orFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' An or operator is Or or OrElse
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression.
    '
    '
    Private Function compiler__orOp_(ByVal objScanned As qbScanner.qbScanner, _
                                     ByRef intIndex As Integer, _
                                     ByVal strSourceCode As String, _
                                     ByRef strOp As String, _
                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "orOp")
        strOp = Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length)
        Select Case UCase(strOp)
            Case "OR"
            Case "ORELSE"
            Case Else : Return compiler__parseFail_("orOp")
        End Select
        compiler__parseEvent_("orOp", _
                                False, _
                                intIndex, 1, _
                                0, 0, _
                                "", _
                                intLevel)
        intIndex += 1       ' I do miss C at times!
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Raise the parse event
    '
    '
    Private Sub compiler__parseEvent_(ByVal strGrammarCategory As String, _
                                        ByVal booTerminal As Boolean, _
                                        ByVal intTokStartIndex As Integer, _
                                        ByVal intTokCount As Integer, _
                                        ByVal intObjectStartIndex As Integer, _
                                        ByVal intObjectCount As Integer, _
                                        ByVal strComment As String, _
                                        ByVal intLevel As Integer)
        Dim strGrammarCategoryWork As String = strGrammarCategory
        If booTerminal Then
            strGrammarCategoryWork = UCase(strGrammarCategory)
        End If
        With OBJstate.usrState.objScanner
            Dim intStartIndex As Integer = Math.Min(intTokStartIndex, .TokenCount)
            Dim intEndIndex As Integer = Math.Min(intStartIndex + intTokCount - 1, .TokenCount)
            raiseEvent_("parseEvent", _
                        strGrammarCategoryWork, _
                        booTerminal, _
                        .QBToken(intStartIndex).StartIndex, _
                        .QBToken(intEndIndex).EndIndex _
                        - _
                        .QBToken(intStartIndex).StartIndex _
                        + _
                        1, _
                        intTokStartIndex, _
                        intTokCount, _
                        intObjectStartIndex, _
                        intObjectCount, _
                        strComment, _
                        intLevel)
        End With
    End Sub

    ' ----------------------------------------------------------------------
    ' Generate the parse fail event and return False
    '
    '
    Private Function compiler__parseFail_(ByVal strGC As String) As Boolean
        raiseEvent_("parseFailEvent", strGC)
        Return False
    End Function

    ' ----------------------------------------------------------------------
    ' powFactor := (+ | -)* term  
    '
    '
    Private Function compiler__powFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "powFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim intMinusCount As Integer
        Do While intIndex <= intEndIndex
            If compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "-", _
                                     intLevel + 1) Then
                intMinusCount += 1
            ElseIf Not compiler__checkToken_(intIndex, _
                                             objScanned, _
                                             strSourceCode, _
                                             intEndIndex, _
                                             "+", _
                                             intLevel + 1) Then
                Exit Do
            End If
        Loop
        Dim booSideEffectsRHS As Boolean
        Dim objConstantValueRHS As qbVariable.qbVariable
        If intMinusCount Mod 2 <> 0 Then
            compiler__genCode_(colPolish, _
                                intIndex, _
                                ENUop.opPushLiteral, _
                                0)
        End If
        If compiler__term_(intIndex, _
                           objScanned, _
                           colPolish, _
                           intEndIndex, _
                           strSourceCode, _
                           objConstantValueRHS, _
                           colVariables, _
                           booSideEffectsRHS, _
                           intLevel + 1) Then
            If intMinusCount Mod 2 <> 0 Then
                objConstantValue = compiler__binaryOpGen_("-", _
                                                          False, _
                                                          Nothing, _
                                                          booSideEffectsRHS, _
                                                          objConstantValueRHS, _
                                                          colPolish, _
                                                          intIndex)
            Else
                objConstantValue = objConstantValueRHS
            End If
            compiler__parseEvent_("powFactor", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("powFactor")
    End Function

    ' ----------------------------------------------------------------------
    ' A power operator is carat or doubled asterisk (they have 
    ' the same precedence, and are the same operator)
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression.
    '
    '
    Private Function compiler__powOp_(ByVal objScanned As qbScanner.qbScanner, _
                                      ByRef intIndex As Integer, _
                                      ByVal strSourceCode As String, _
                                      ByRef strOp As String, _
                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "powOp")
        strOp = Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length)
        Select Case strOp
            Case "^"
            Case "**"
            Case Else : Return compiler__parseFail_("powOp")
        End Select
        compiler__parseEvent_("powOp", _
                                False, _
                                intIndex, 1, _
                                0, 0, _
                                "", _
                                intLevel)
        intIndex += 1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Make a new entry on the compiler's nesting stack
    '
    '
    Private Overloads Function compiler__pushNesting_(ByRef objNesting As Stack, _
                                                      ByVal strStartLabel As String, _
                                                      ByVal enuNestType As ENUnesting, _
                                                      ByVal objInfo As Object, _
                                                      ByVal intLoopExit As Integer, _
                                                      ByVal intLine As Integer, _
                                                      ByVal strStackImage As String) As Boolean
        Dim usrNesting As TYPnesting
        With usrNesting
            .intLine = intLine
            .strStartLabel = strStartLabel
            .enuNestType = enuNestType
            .objInfo = objInfo
            .strStackImage = strStackImage
            If intLoopExit >= 0 Then
                Try
                    ReDim .intLoopExits(CInt(IIf(intLoopExit = 0, 0, 1)))
                Catch objException As Exception
                    errorHandler_("Internal compiler error: " & _
                                  "cannot allocate loop exit table: " & _
                                  Err.Number & " " & Err.Description, _
                                  "compiler__pushNesting_", _
                                  "Making object unusable and returning False", _
                                  objException)
                    OBJstate.usrState.booUsable = False : Return compiler__parseFail_("pushNesting")
                End Try
                If intLoopExit > 0 Then .intLoopExits(1) = intLoopExit
            End If
        End With
        Try
            objNesting.Push(usrNesting)
        Catch objException As Exception
            errorHandler_("Internal compiler error: cannot extend the nesting stack: " & _
                          Err.Number & " " & Err.Description, _
                          "compiler__pushNesting_", _
                          "Making object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False : Return compiler__parseFail_("pushNesting")
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' relOp := "<" | ">" | "=" | "<=" | ">=" | "=" | "<>"	
    '
    '
    ' Note that this method on success increments the index to the scanned
    ' expression, and returns the operator in a reference parameter.
    '
    '
    Private Function compiler__relOp_(ByVal objScanned As qbScanner.qbScanner, _
                                      ByRef intIndex As Integer, _
                                      ByVal strSourceCode As String, _
                                      ByRef strOp As String, _
                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "relOp")
        strOp = Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length)
        Select Case strOp
            Case "<"
            Case ">"
            Case "="
            Case "<>"
            Case ">="
            Case "<="
            Case Else : Return compiler__parseFail_("relOp")
        End Select
        compiler__parseEvent_("relOp", _
                                False, _
                                intIndex, 1, _
                                0, 0, _
                                "", _
                                intLevel)
        intIndex += 1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' relFactor := addFactor [relFactorRHS]
    '
    '
    Private Function compiler__relFactor_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objConstantValue As qbVariable.qbVariable, _
                                          ByRef colVariables As Collection, _
                                          ByRef booSideEffects As Boolean, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "relFactor")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__addFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValue, _
                                    colVariables, _
                                    booSideEffects, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("relFactor")
        End If
        If intIndex > intEndIndex Then
            compiler__parseEvent_("relFactor", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "No RHS", _
                                    intLevel)
            Return (True)
        End If
        Dim intIndex2 As Integer = intIndex
        If Not compiler__relFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("relFactor")
        compiler__parseEvent_("relFactor", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "Has RHS", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' relFactorRHS = addOp addFactor [relFactorRHS]
    '
    '
    Private Function compiler__relFactorRHS_(ByRef intIndex As Integer, _
                                             ByVal objScanned As qbScanner.qbScanner, _
                                             ByRef colPolish As Collection, _
                                             ByVal intEndIndex As Integer, _
                                             ByVal strSourceCode As String, _
                                             ByRef objConstantValue As qbVariable.qbVariable, _
                                             ByRef colVariables As Collection, _
                                             ByRef booSideEffects As Boolean, _
                                             ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "relFactorRHS")
        If intIndex > intEndIndex Then Return compiler__parseFail_("relFactorRHS")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim strAddOp As String
        If Not compiler__addOp_(objScanned, intIndex, strSourceCode, strAddOp, intLevel + 1) Then
            Return compiler__parseFail_("relFactorRHS")
        End If
        Dim booSideEffectsRHS As Boolean
        Dim objConstantValueRHS As qbVariable.qbVariable
        If Not compiler__addFactor_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objConstantValueRHS, _
                                    colVariables, _
                                    booSideEffectsRHS, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("relFactorRHS")
        End If
        objConstantValue = compiler__binaryOpGen_(strAddOp, _
                                                  booSideEffects, _
                                                  objConstantValue, _
                                                  booSideEffectsRHS, _
                                                  objConstantValueRHS, _
                                                  colPolish, _
                                                  intIndex)
        Dim intIndex2 As Integer = intIndex
        If Not compiler__relFactorRHS_(intIndex, _
                                          objScanned, _
                                          colPolish, _
                                          intEndIndex, _
                                          strSourceCode, _
                                          objConstantValue, _
                                          colVariables, _
                                          booSideEffects, _
                                          intLevel + 1) _
            AndAlso _
            intIndex2 <> intIndex Then Return compiler__parseFail_("relFactorRHS")
        compiler__parseEvent_("relFactorRHS", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' singleImmediateCommand := explicitAssignment | expression 
    '
    '
    Private Function compiler__singleImmediateCommand_(ByRef intIndex As Integer, _
                                                       ByVal objScanned As qbScanner.qbScanner, _
                                                       ByRef colPolish As Collection, _
                                                       ByVal intEndIndex As Integer, _
                                                       ByVal strSourceCode As String, _
                                                       ByRef colVariables As Collection, _
                                                       ByRef objImmediateValue As qbVariable.qbVariable, _
                                                       ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "singleImmediateCommand")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        objImmediateValue = Nothing
        If compiler__explicitAssignment_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         colVariables, _
                                         intLevel + 1) Then
            compiler__parseEvent_("singleImmediateCommand", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "Expression", _
                                    intLevel)
            Return (True)
        End If
        intIndex = intIndex1
        If compiler__expression_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 objImmediateValue, _
                                 colVariables, _
                                 intLevel + 1) _
           AndAlso _
           intIndex > objScanned.TokenCount Then
            compiler__parseEvent_("singleImmediateCommand", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "Expression", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("singleImmediateCommand")
    End Function

    ' ----------------------------------------------------------------------
    ' sourceProgram := optionStmt  
    ' sourceProgram := sourceProgramBody
    ' sourceProgram := optionStmt logicalNewline sourceProgramBody
    '
    '
    Private Function compiler__sourceProgram_(ByRef intIndex As Integer, _
                                              ByVal objScanned As qbScanner.qbScanner, _
                                              ByRef colPolish As Collection, _
                                              ByVal intEndIndex As Integer, _
                                              ByVal strSourceCode As String, _
                                              ByRef colVariables As Collection, _
                                              ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "sourceProgram")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__optionStmt_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intLevel + 1) Then
            If intIndex > intEndIndex Then Return (True)
            If Not compiler__logicalNewline_(intIndex, _
                                             objScanned, _
                                             strSourceCode, _
                                             intEndIndex, _
                                             intLevel + 1) Then
                compiler__errorHandler_("Option statement is not terminated by a new line", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("sourceProgram")
            End If
        End If
        Dim objNestingStack As New Stack           ' Of TYPnesting only
        If Not compiler__pushNesting_(objNestingStack, _
                                      "", _
                                      ENUnesting.openCode, _
                                      0, _
                                      -1, _
                                      objScanned.QBToken(intIndex).LineNumber, _
                                      "") Then Return compiler__parseFail_("sourceProgram")
        If Not compiler__sourceProgramBody_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            intEndIndex, _
                                            strSourceCode, _
                                            colVariables, _
                                            objNestingStack, _
                                            intLevel + 1) Then Return compiler__parseFail_("sourceProgram")
        If objNestingStack.Count <> 1 Then
            compiler__errorHandler_("Source code contains " & objNestingStack.Count & " " & _
                                    "unbalanced For/Do/If statement(s) at " & _
                                    compiler__sourceProgramBody__unbalList_(objNestingStack), _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("sourceProgram")
        End If
        compiler__parseEvent_("sourceProgram", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "Expression", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' sourceProgramBody := ( openCode | moduleDefinition ) +
    '
    '
    Private Function compiler__sourceProgramBody_(ByRef intIndex As Integer, _
                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                  ByRef colPolish As Collection, _
                                                  ByVal intEndIndex As Integer, _
                                                  ByVal strSourceCode As String, _
                                                  ByRef colVariables As Collection, _
                                                  ByRef objNestingStack As Stack, _
                                                  ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "sourceProgramBody")
        Dim booReturn As Boolean
        Dim intIndex1 As Integer = intIndex
        Dim intIndex2 As Integer = colPolish.Count
        Do While intIndex <= intEndIndex _
                 AndAlso _
                 compiler__openCode_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     colVariables, _
                                     objNestingStack, _
                                     intLevel + 1) _
                  OrElse _
                  compiler__moduleDefinition_(intIndex, _
                                              objScanned, _
                                              colPolish, _
                                              intEndIndex, _
                                              strSourceCode, _
                                              colVariables, _
                                              objNestingStack, _
                                              intLevel + 1)
            booReturn = True
            compiler__parseEvent_("sourceProgramBody", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intIndex2, _
                                    colPolish.Count - intIndex2, _
                                    "", _
                                    intLevel)
        Loop
        If Not booReturn Then Return compiler__parseFail_("sourceProgramBody")
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' sourceProgram recursion
    '
    '
    Private Function compiler__sourceProgramBody__(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef objNestingStack As Stack, _
                                                   ByRef colVariables As Collection, _
                                                   ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex
        If Not compiler__statement_(intIndex, _
                                    objScanned, _
                                    colPolish, _
                                    intEndIndex, _
                                    strSourceCode, _
                                    objNestingStack, _
                                    colVariables, _
                                    intLevel + 1) Then
            Return compiler__parseFail_("sourceProgramBody")
        End If
        If Not compiler__logicalNewline_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         intLevel + 1) Then
            Return (True)
        End If
        If compiler__sourceProgramBody__(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         objNestingStack, _
                                         colVariables, _
                                         intLevel) Then
            Return (True)
        End If
        Return compiler__parseFail_("sourceProgramBody")
    End Function

    ' ----------------------------------------------------------------------
    ' Returns the comma-separated list of unbalanced statements
    '
    '
    Private Function compiler__sourceProgramBody__unbalList_(ByVal objNesting As Stack) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As New Text.StringBuilder
        Dim usrNesting As TYPnesting
        With objNesting
            For intIndex1 = .Count To 1 Step -1
                _OBJutilities.append(objStringBuilder, ", ", compiler__nesting2String_(CType(objNesting.Pop, TYPnesting)))
            Next intIndex1
        End With
        Return objStringBuilder.ToString
    End Function

    ' ----------------------------------------------------------------------
    ' statement := [integer | identifier:] statementBody
    '
    '
    Private Function compiler__statement_(ByRef intIndex As Integer, _
                                          ByVal objScanned As qbScanner.qbScanner, _
                                          ByRef colPolish As Collection, _
                                          ByVal intEndIndex As Integer, _
                                          ByVal strSourceCode As String, _
                                          ByRef objNestingStack As Stack, _
                                          ByRef colVariables As Collection, _
                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "statement")
        Dim strLabel As String
        If Not compiler__genCode_(colPolish, intIndex, ENUop.opRem, 0, "") Then
            Return compiler__parseFail_("statement")
        End If
        Dim intIndex1 As Integer = intIndex
        Dim intIndex2 As Integer = colPolish.Count
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger, _
                                 intLevel + 1) Then
            strLabel = Mid(strSourceCode, _
                           objScanned.QBToken(intIndex - 1).StartIndex, _
                           objScanned.QBToken(intIndex - 1).Length)
        Else
            If compiler__checkToken_(intIndex, _
                                        objScanned, _
                                        strSourceCode, _
                                        intEndIndex, _
                                        qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                        intLevel + 1) _
               AndAlso _
               compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon, _
                                     intLevel + 1) Then
                strLabel = Mid(strSourceCode, _
                               objScanned.QBToken(intIndex - 2).StartIndex, _
                               objScanned.QBToken(intIndex - 2).Length)
            Else
                intIndex = intIndex1
            End If
        End If
        If strLabel <> "" Then
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel) Then
                Return compiler__parseFail_("statement")
            End If
        End If
        If Not compiler__statementBody_(intIndex, _
                                        objScanned, _
                                        colPolish, _
                                        intEndIndex, _
                                        strSourceCode, _
                                        objNestingStack, _
                                        colVariables, _
                                        intLevel) Then
            Return compiler__parseFail_("statement")
        End If
        If intIndex2 <> 0 Then
            Dim objPolishHandle As qbPolish.qbPolish = _
                CType(colPolish.Item(intIndex2), qbPolish.qbPolish)
            objPolishHandle.Comment = _
                    "***** " & _
                    rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1)
            raiseEvent_("codeChangeEvent", objPolishHandle, intIndex2)
        End If
        compiler__parseEvent_("sourceProgramBody", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intIndex2, _
                                colPolish.Count - intIndex2, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' statementBody := ctlStatementBody | 
    '                  unconditionalStatementBody |
    '                  assignmentStmt |
    '                  functionCall  
    '
    '
    Private Function compiler__statementBody_(ByRef intIndex As Integer, _
                                              ByVal objScanned As qbScanner.qbScanner, _
                                              ByRef colPolish As Collection, _
                                              ByVal intEndIndex As Integer, _
                                              ByVal strSourceCode As String, _
                                              ByRef objNestingStack As Stack, _
                                              ByRef colVariables As Collection, _
                                              ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "statementBody")
        Dim booReturn As Boolean
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__statementBody__unconditional_(intIndex, _
                                                   objScanned, _
                                                   colPolish, _
                                                   intEndIndex, _
                                                   strSourceCode, _
                                                   colVariables, _
                                                   objNestingStack, _
                                                   intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__ctl_(intIndex, _
                                                objScanned, _
                                                colPolish, _
                                                intEndIndex, _
                                                strSourceCode, _
                                                objNestingStack, _
                                                colVariables, _
                                                intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__assignmentStmt_(intIndex, _
                                                        objScanned, _
                                                        colPolish, _
                                                        intEndIndex, _
                                                        strSourceCode, _
                                                        colVariables, _
                                                        intLevel + 1) Then
            booReturn = True
        ElseIf compiler__functionCall_(intIndex, _
                                        objScanned, _
                                        colPolish, _
                                        intEndIndex, _
                                        strSourceCode, _
                                        colVariables, _
                                        intLevel + 1) Then
            booReturn = True
        End If
        If Not booReturn Then
            compiler__errorHandler_("Statement is unrecognizable", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("statementBody")
        End If
        compiler__parseEvent_("statementBody", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' ctlStatementBody := defFnr | dim | doHeader | else | endif | forHeader | 
    '                     forNext | if | whileHeader | loopOrWend 
    '
    '
    Private Function compiler__statementBody__ctl_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef objNestingStack As Stack, _
                                                   ByRef colVariables As Collection, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "ctlStatementBody")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__statementBody__dim_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__doHeader_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, colVariables, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__else_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__endif_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__forHeader_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, colVariables, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__forNext_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__if_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, colVariables, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__whileHeader_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, colVariables, intLevel + 1) Then
            Return (True)
        ElseIf compiler__statementBody__loopOrWend_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, colVariables, intLevel + 1) Then
            Return (True)
        Else
            Return compiler__parseFail_("ctlStatementBody")
        End If
        compiler__parseEvent_("ctlStatementBody", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
    End Function

    ' ----------------------------------------------------------------------
    ' assignmentStmt := explicitAssignment | implicitAssignment
    '
    '
    Private Function compiler__statementBody__assignmentStmt_(ByRef intIndex As Integer, _
                                                              ByVal objScanned As qbScanner.qbScanner, _
                                                              ByRef colPolish As Collection, _
                                                              ByVal intEndIndex As Integer, _
                                                              ByVal strSourceCode As String, _
                                                              ByRef colVariables As Collection, _
                                                              ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "assignmentStmt")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__explicitAssignment_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         colVariables, _
                                         intLevel + 1) _
           OrElse _
           compiler__implicitAssignment_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         colVariables, _
                                         intLevel + 1) Then
            compiler__parseEvent_("statementBody", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("assignmentStmt")
    End Function

    ' ----------------------------------------------------------------------
    ' circle := Circle ( expression comma expression ) comma expression
    '
    '
    ' Note that at this writing we do not support the extended syntax
    '
    '
    Private Function compiler__statementBody__circle_(ByRef intIndex As Integer, _
                                                      ByVal objScanned As qbScanner.qbScanner, _
                                                      ByRef colPolish As Collection, _
                                                      ByVal intEndIndex As Integer, _
                                                      ByVal strSourceCode As String, _
                                                      ByRef colVariables As Collection, _
                                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "circle")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "Circle", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("circle")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "(", _
                                     intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("circle")
        End If
        Dim intParms As Integer
        If Not compiler__expressionList_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         2, _
                                         2, _
                                         colVariables, _
                                         intParms, _
                                         intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("circle")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     ")", _
                                     intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("circle")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma, _
                                     intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("circle")
        End If
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     colVariables, _
                                     intLevel + 1) Then
            intIndex = intIndex1 : Return compiler__parseFail_("circle")
        End If
        If Not compiler__genCode_(colPolish, intIndex, ENUop.opCircle, 0, _
                                  "Draw a circle") Then
            Return compiler__parseFail_("circle")
        End If
        compiler__parseEvent_("circle", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' cls = "CLS"
    '
    '
    Private Function compiler__statementBody__cls_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "cls")
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "CLS", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("cls")
        End If
        If Not compiler__genCode_(colPolish, _
                                  intIndex, _
                                  ENUop.opCls, _
                                  Nothing, _
                                  "Clear the screen") Then
            Return compiler__parseFail_("cls")
        End If
        compiler__parseEvent_("cls", _
                                False, _
                                intIndex - 1, _
                                1, _
                                colPolish.Count - 1, _
                                1, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Comment := REM NONEWLINE
    ' Comment := APOSTROPHE NONEWLINE
    ' Comment := EMPTYLINE
    '
    '
    ' Note that "NONEWLINE" means a string of characters that does not include
    ' a physical newline.  Colons do not break lines after the REMark. The
    ' intIndex is left at the new line.
    '
    '
    Private Function compiler__statementBody__comment_(ByRef intIndex As Integer, _
                                                       ByVal objScanned As qbScanner.qbScanner, _
                                                       ByRef colPolish As Collection, _
                                                       ByVal intEndIndex As Integer, _
                                                       ByVal strSourceCode As String, _
                                                       ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "comment")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        Dim intIndexLineEnd As Integer = compiler__findNewline_(intIndex, _
                                                                objScanned, _
                                                                intEndIndex, _
                                                                False, _
                                                                booColonIsNewLine:=False)
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "REM", _
                                     intLevel + 1) _
           AndAlso _
           Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeApostrophe, _
                                     intLevel + 1) _
           AndAlso _
           intIndexLineEnd <> intIndex Then
            Return compiler__parseFail_("comment")
        End If
        compiler__parseEvent_("comment", _
                                False, _
                                intIndex1, _
                                intIndexLineEnd - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        intIndex = intIndexLineEnd
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Data := Data constantList
    '
    '
    Private Function compiler__statementBody__data_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "data")
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "DATA", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("data")
        End If
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = compiler__findNewline_(intIndex, objScanned, intEndIndex, False) - 1
        Dim intStartIndex As Integer = intIndex1
        Dim objConstantValue As qbVariable.qbVariable
        Dim booConstantFoldingSave As Boolean
        Do
            booConstantFoldingSave = OBJstate.usrState.booConstantFolding
            OBJstate.usrState.booConstantFolding = True
            If Not compiler__expression_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intIndex1, _
                                         strSourceCode, _
                                         objConstantValue, _
                                         Nothing, _
                                         intLevel + 1) Then
                OBJstate.usrState.booConstantFolding = booConstantFoldingSave
                If intIndex1 <> intStartIndex Then
                    compiler__errorHandler_("Invalid Data statement", _
                                            intIndex, _
                                            objScanned, _
                                            strSourceCode)
                    Return compiler__parseFail_("data")
                End If
                Exit Do
            End If
            OBJstate.usrState.booConstantFolding = booConstantFoldingSave
            If (objConstantValue Is Nothing) Then
                compiler__errorHandler_("Data statements can contain only constants", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("data")
            End If
            Try
                OBJstate.usrState.colReadData.Add(objConstantValue)
            Catch
                compiler__errorHandler_("Cannot update the Read Data queue during compilation: details: " & _
                                        Err.Number & " " & Err.Description, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("data")
            End Try
            If Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intIndex1, _
                                         qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma, _
                                         intLevel + 1) Then Exit Do
        Loop
        compiler__parseEvent_("data", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' dim := Dim lValueList
    '
    '
    Private Function compiler__statementBody__dim_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef colVariables As Collection, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "dim")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "DIM", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("dim")
        End If
        If Not compiler__lValueList_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     ENUlValueContext.definition, _
                                     colVariables, _
                                     intLevel + 1) Then Return compiler__parseFail_("dim")
        compiler__parseEvent_("dim", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' doHeader := Do [ doCondition ] 
    '
    '
    Private Function compiler__statementBody__doHeader_(ByRef intIndex As Integer, _
                                                        ByVal objScanned As qbScanner.qbScanner, _
                                                        ByRef colPolish As Collection, _
                                                        ByVal intEndIndex As Integer, _
                                                        ByVal strSourceCode As String, _
                                                        ByRef objNestingStack As Stack, _
                                                        ByRef colVariables As Collection, _
                                                        ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "doHeader")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndexStart As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "DO", _
                                     intLevel + 1) Then Return compiler__parseFail_("doHeader")
        Dim strLabel As String = compiler__getLabel_()
        If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "") Then Return compiler__parseFail_("doHeader")
        Dim intIndex1 As Integer = colPolish.Count - 1
        Dim booSave As Boolean = OBJstate.usrState.booConstantFolding
        OBJstate.usrState.booConstantFolding = False
        Dim booHasCondition As Boolean = compiler__whileUntilClause_(intIndex, _
                                                                     objScanned, _
                                                                     colPolish, _
                                                                     intEndIndex, _
                                                                     strSourceCode, _
                                                                     objNestingStack, _
                                                                     colVariables, _
                                                                     intLevel + 1)
        OBJstate.usrState.booConstantFolding = booSave
        CType(colPolish(intIndex1), qbPolish.qbPolish).Comment = _
            "Do " & rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1)
        If compiler__pushNesting_(objNestingStack, strLabel, ENUnesting.doLoop, Nothing, _
                                  CInt(IIf(booHasCondition, 1, 0)), _
                                  objScanned.QBToken(intIndex).LineNumber, _
                                  "") Then
            compiler__parseEvent_("doHeader", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("doHeader")
    End Function


    ' ----------------------------------------------------------------------
    ' else := Else 
    '
    '
    Private Function compiler__statementBody__else_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByRef objNestingStack As Stack, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "else")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "ELSE", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("else")
        End If
        Dim usrNesting As TYPnesting = CType(objNestingStack.Peek, TYPnesting)
        With usrNesting
            If .enuNestType <> ENUnesting.ifThen Then
                compiler__errorHandler_("Else clause is not balanced with an If statement", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("else")
            End If
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opJump, 0, "Jump around the Else clause") Then
                Return compiler__parseFail_("else")
            End If
            Dim strLabel As String = compiler__getLabel_()
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "Else clause") Then Return compiler__parseFail_("else")
            CType(colPolish.Item(CInt(.objInfo)), qbPolish.qbPolish).Operand = strLabel
            objNestingStack.Pop()
            If Not compiler__pushNesting_(objNestingStack, _
                                          "", _
                                          ENUnesting.elseClause, _
                                          colPolish.Count - 1, _
                                          -1, _
                                          objScanned.QBToken(intIndex).LineNumber, _
                                          "") Then
                Return compiler__parseFail_("else")
            End If
            compiler__parseEvent_("else", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        compiler__parseFail_("else")
    End Function

    ' ----------------------------------------------------------------------
    ' end = End (followed immediately by newline: this recognizer does not
    '            increment past the newline, only checks for it)
    '
    '
    Private Function compiler__statementBody__end_(ByRef intIndex As Integer, _
                                                   ByVal objScanned As qbScanner.qbScanner, _
                                                   ByRef colPolish As Collection, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal strSourceCode As String, _
                                                   ByRef objNestingStack As Stack, _
                                                   ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "end")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "END", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("end")
        End If
        If compiler__findNewline_(intIndex, objScanned, intEndIndex, False) <> intIndex Then
            intIndex = intIndex1
            Return compiler__parseFail_("end")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opEnd, 0, "End of processing") Then
            compiler__parseEvent_("end", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("end")
    End Function

    ' ----------------------------------------------------------------------
    ' endif = EndIf | End If
    '
    '
    Private Function compiler__statementBody__endif_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef objNestingStack As Stack, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "endif")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "END", _
                                 intLevel + 1) Then
            If Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         "IF", _
                                         intLevel + 1) Then
                intIndex = intIndex1 : Return compiler__parseFail_("endIf")
            End If
        ElseIf Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         "ENDIF", _
                                         intLevel + 1) Then
            Return compiler__parseFail_("endIf")
        End If
        Dim usrNesting As TYPnesting = CType(objNestingStack.Peek, TYPnesting)
        With usrNesting
            If .enuNestType <> ENUnesting.ifThen AndAlso .enuNestType <> ENUnesting.elseClause Then
                compiler__errorHandler_("End If is not balanced with an If or an Else statement", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("endIf")
            End If
            Dim strLabel As String = compiler__getLabel_()
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "End If") Then Return compiler__parseFail_("endIf")
            Dim intIndex2 As Integer = CInt(.objInfo)
            Dim objPolish As qbPolish.qbPolish = CType(colPolish.Item(intIndex2), qbPolish.qbPolish)
            objPolish.Operand = strLabel
            raiseEvent_("codeChangeEvent", objPolish, intIndex2)
            objNestingStack.Pop()
            compiler__parseEvent_("endif", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        compiler__parseFail_("endif")
    End Function

    ' ----------------------------------------------------------------------
    ' exit := Exit [ Do | For | While | Function | Sub ] 
    '
    '
    ' Note: as of this writing (2-2-2002) the keywords Function and Sub are
    ' supported in syntax but will cause an error, since at this time, 
    ' functions and subroutines haven't been implemented.
    '
    '
    Private Function compiler__statementBody__exit_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByRef objNestingStack As Stack, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "exit")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "EXIT", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("exit")
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "DO", _
                                 intLevel + 1) Then
            If compiler__statementBody__exit__ctlStructure_(intIndex, _
                                                            objScanned, _
                                                            colPolish, _
                                                            strSourceCode, _
                                                            ENUnesting.doLoop, _
                                                            objNestingStack) Then
                compiler__parseEvent_("statementBody", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "Exit Do", _
                                        intLevel)
                Return (True)
            End If
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "FOR", _
                                 intLevel + 1) Then
            If compiler__statementBody__exit__ctlStructure_(intIndex, _
                                                                objScanned, _
                                                                colPolish, _
                                                                strSourceCode, _
                                                                ENUnesting.forLoop, _
                                                                objNestingStack) Then
                compiler__parseEvent_("statementBody", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "Exit For", _
                                        intLevel)
                Return (True)
            End If
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "WHILE", _
                                 intLevel + 1) Then
            If compiler__statementBody__exit__ctlStructure_(intIndex, _
                                                                objScanned, _
                                                                colPolish, _
                                                                strSourceCode, _
                                                                ENUnesting.whileLoop, _
                                                                objNestingStack) Then
                compiler__parseEvent_("statementBody", _
                                        False, _
                                        intIndex1, _
                                        intIndex - intIndex1, _
                                        intCountPrevious + 1, _
                                        colPolish.Count - intCountPrevious, _
                                        "Exit While", _
                                        intLevel)
                Return (True)
            End If
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "FUNCTION", _
                                 intLevel + 1) Then
            compiler__errorHandler_("Exit Function used: functions aren't supported yet: Stop statement generated.", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "SUB", _
                                 intLevel + 1) Then
            compiler__errorHandler_("Exit Sub used: subroutines aren't supported yet: Stop statement generated.", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opEnd, Nothing, "Terminate program") Then
            compiler__parseEvent_("exit", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "Exit program", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("exit")
    End Function

    ' ----------------------------------------------------------------------
    ' Compile a control structure exit: generate goto and update the list
    '
    '
    Private Function compiler__statementBody__exit__ctlStructure_(ByVal intIndex As Integer, _
                                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                                  ByRef colPolish As Collection, _
                                                                  ByVal strSourceCode As String, _
                                                                  ByVal enuNestingType As ENUnesting, _
                                                                  ByRef objNestingStack As Stack) As Boolean
        Dim colSaveStack As New Collection
        Dim intIndex1 As Integer
        Dim strNest As String = compiler__nestingType2String_(enuNestingType)
        With colSaveStack
            Dim booFound As Boolean
            Do While objNestingStack.Count > 0
                If CType(objNestingStack.Peek, TYPnesting).enuNestType = enuNestingType Then
                    booFound = True : Exit Do
                End If
                .Add(objNestingStack.Pop)
            Loop
            If Not booFound Then
                For intIndex1 = .Count To 1 Step -1
                    objNestingStack.Push(.Item(intIndex1))
                Next intIndex1
                compiler__errorHandler_("Exit type " & _OBJutilities.enquote(strNest) & " " & _
                                        "cannot be found in the surrounding code", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return (False)
            End If
        End With
        If Not compiler__genCode_(colPolish, intIndex, ENUop.opJump, 0, "Exit " & strNest & " structure") Then
            For intIndex1 = colSaveStack.Count To 1 Step -1
                objNestingStack.Push(colSaveStack.Item(intIndex1))
            Next intIndex1
            Return (False)
        End If
        Dim usrNestingStack As TYPnesting = CType(objNestingStack.Pop, TYPnesting)
        With usrNestingStack
            Try
                ReDim Preserve .intLoopExits(UBound(.intLoopExits) + 1)
            Catch
                compiler__errorHandler_("Too many loop exits defined: Redim error: " & Err.Number & " " & Err.Description, _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                objNestingStack.Push(usrNestingStack)
                For intIndex1 = colSaveStack.Count To 1 Step -1
                    objNestingStack.Push(colSaveStack.Item(intIndex1))
                Next intIndex1
                Return (False)
            End Try
            .intLoopExits(UBound(.intLoopExits)) = colPolish.Count
            objNestingStack.Push(usrNestingStack)
            For intIndex1 = colSaveStack.Count To 1 Step -1
                objNestingStack.Push(colSaveStack.Item(intIndex1))
            Next intIndex1
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' forHeader := For lValue = expression To expression [ Step expression ] 
    '
    '
    ' Note that for For loops and at runtime, the following is the For stack
    ' frame:
    '
    '
    '      Control Variable location
    '      Step value 
    '      Final value
    '
    '
    Private Function compiler__statementBody__forHeader_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByRef objNestingStack As Stack, _
                                                         ByRef colVariables As Collection, _
                                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "forHeader")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndexStart As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "FOR", _
                                     intLevel + 1) Then Return compiler__parseFail_("forHeader")
        Dim intCtlVarMemReference As Integer
        Dim intIndex1 As Integer = intIndex
        Dim objConstantEndValue As qbVariable.qbVariable
        Dim objConstantStartValue As qbVariable.qbVariable
        Dim objConstantStepValue As qbVariable.qbVariable
        ' --- Parse control lValue followed by the equals sign
        If Not compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intCtlVarMemReference, _
                                 colVariables, _
                                 False, _
                                 intLevel + 1) Then
            compiler__errorHandler_("For keyword is not followed by the For control variable", _
                                    intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        If intCtlVarMemReference < 0 Then
            compiler__errorHandler_("For loop requires a control variable", _
                                    intIndex1, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        Dim strCtlLValue As String = rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1)
        Dim strStackImage As String
        If Not compiler__genCode_(colPolish, intIndex, _
                                  ENUop.opPushLiteral, _
                                  intCtlVarMemReference, _
                                  "Push the control variable " & strCtlLValue) Then
            Return compiler__parseFail_("forHeader")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "=", _
                                     intLevel + 1) Then
            compiler__errorHandler_("For control variable is not followed by an equals sign", _
                                    intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        ' --- Parse and generate starting value
        Dim strCode As String
        intIndex1 = intIndex
        Dim intIndex2 As Integer = colPolish.Count + 1
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantStartValue, _
                                     colVariables, _
                                     intLevel + 1) Then
            compiler__errorHandler_("For starting value or expression is not present or not valid", _
                                    intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        If (objConstantStartValue Is Nothing) Then
            CType(colPolish(intIndex2), qbPolish.qbPolish).Comment = _
                "ctlVariable->ctlVariable,initialValue"
        Else
            If Not compiler__genCode_(colPolish, intIndex, _
                                      ENUop.opPushLiteral, _
                                      objConstantStartValue, _
                                      "indexLoc" & strCtlLValue & _
                                      "ctlVariable->ctlVariable,initialValue") Then
                Return compiler__parseFail_("forHeader")
            End If
        End If
        If Not compiler__genCode_(colPolish, intIndex, _
                                  ENUop.opPopIndirect, _
                                  Nothing, _
                                  "ctlVariable,initialValue->ctlVariable") Then
            Return compiler__parseFail_("forHeader")
        End If
        ' --- Parse and generate ending value
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "TO", _
                                     intLevel + 1) Then
            compiler__errorHandler_("In For loop header, no To keyword found", intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        intIndex1 = intIndex
        intIndex2 = colPolish.Count + 1
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     objConstantEndValue, _
                                     colVariables, _
                                     intLevel + 1) Then
            compiler__errorHandler_("For terminating value or expression is not present or not valid", _
                                    intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("forHeader")
        End If
        If (objConstantEndValue Is Nothing) Then
            CType(colPolish.Item(intIndex2), qbPolish.qbPolish).Comment = _
                "ctlVariable->ctlVariable,finalValue"
        Else
            If Not compiler__genCode_(colPolish, intIndex, _
                                      ENUop.opPushLiteral, _
                                      objConstantEndValue, _
                                      "ctlVariable->ctlVariable,finalValue") Then Return compiler__parseFail_("forHeader")
        End If
        ' --- Parse and generate optional step value
        If Not compiler__genCode_(colPolish, intIndex, _
                                    ENUop.opRotate, _
                                    1, _
                                    "ctlVariable,finalValue->" & _
                                    "finalValue,ctlVariable") Then
            Return compiler__parseFail_("forHeader")
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "STEP", _
                                 intLevel + 1) Then
            intIndex1 = intIndex
            intIndex2 = colPolish.Count + 1
            If Not compiler__expression_(intIndex, _
                                         objScanned, _
                                         colPolish, _
                                         intEndIndex, _
                                         strSourceCode, _
                                         objConstantStepValue, _
                                         colVariables, _
                                         intLevel + 1) Then
                compiler__errorHandler_("For Step keyword is not followed by a Step value or expression", _
                                        intIndex, objScanned, strSourceCode)
                Return compiler__parseFail_("forHeader")
            End If
        Else
            objConstantStepValue = _OBJqbVariable.mkVariableFromValue("1")
        End If
        If (objConstantStepValue Is Nothing) Then
            CType(colPolish(intIndex2), qbPolish.qbPolish).Comment = _
                  "ctlVariable,initialValue->ctlVariable,finalValue,stepValue"
        Else
            If Not compiler__genCode_(colPolish, intIndex, _
                                      ENUop.opPushLiteral, _
                                      objConstantStepValue, _
                                      "finalValue,ctlVariable->" & _
                                      "finalValue,ctlVariable,stepValue") Then
                Return compiler__parseFail_("forHeader")
            End If
        End If
        If Not compiler__genCode_(colPolish, intIndex, _
                                    ENUop.opRotate, _
                                    1, _
                                    "finalValue,ctlVariable,stepValue->" & _
                                    "finalValue,stepValue,ctlVariable") Then
            Return compiler__parseFail_("forHeader")
        End If
        ' --- Generate the loop label
        Dim strLabel As String = compiler__getLabel_()
        If Not compiler__genCode_(colPolish, intIndex, _
                                  ENUop.opLabel, _
                                  strLabel, _
                                  "For loop starts here") Then Return compiler__parseFail_("forHeader")
        ' --- Generate the For test code with an empty label: it will be filled in later
        If Not compiler__genCode_(colPolish, intIndex, _
                                  ENUop.opForTest, _
                                  Nothing, _
                                  "Test For condition using the stack frame") Then Return compiler__parseFail_("forHeader")
        ' --- Update the nesting stack
        If compiler__pushNesting_(objNestingStack, _
                                  strLabel, _
                                  ENUnesting.forLoop, _
                                  strCtlLValue, _
                                  colPolish.Count, _
                                  objScanned.QBToken(intIndex1).LineNumber, _
                                  strStackImage) Then
            compiler__parseEvent_("forHeader", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("forHeader")
    End Function

    ' ----------------------------------------------------------------------
    ' forNext := Next identifier 
    '
    '
    Private Function compiler__statementBody__forNext_(ByRef intIndex As Integer, _
                                                       ByVal objScanned As qbScanner.qbScanner, _
                                                       ByRef colPolish As Collection, _
                                                       ByVal intEndIndex As Integer, _
                                                       ByVal strSourceCode As String, _
                                                       ByRef objNestingStack As Stack, _
                                                       ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "forNext")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "NEXT", _
                                     intLevel + 1) Then Return compiler__parseFail_("forNext")
        If objNestingStack.Count < 1 Then
            compiler__errorHandler_("There is nothing in the nesting stack, " & _
                                    "indicating a probable compiler error and/or a possible error in the source code", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("forNext")
        End If
        Dim usrNesting As TYPnesting = CType(objNestingStack.Pop, TYPnesting)
        If usrNesting.enuNestType <> ENUnesting.forLoop Then
            compiler__errorHandler_("Next is not matched with a For", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("forNext")
        End If
        Dim strForLValue As String = CStr(usrNesting.objInfo)
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                 intLevel + 1) Then
            With objScanned.QBToken(intIndex - 1)
                Dim strNextIdentifier As String = Mid(strSourceCode, .StartIndex, .Length)
                If UCase(strNextIdentifier) <> UCase(strForLValue) Then
                    compiler__errorHandler_("For control variable " & strForLValue & " " & _
                                            "does not match the variable " & strNextIdentifier & " " & _
                                            "in the balanced Next statement: " & _
                                            "assuming usability of For control variable", _
                                            intIndex, objScanned, strSourceCode)
                End If
            End With
        Else
            compiler__errorHandler_("For control variable " & strForLValue & " " & _
                                    "not found in balanced Next statement: " & _
                                    "assuming usability of For control variable " & strForLValue, _
                                    intIndex, objScanned, strSourceCode)
        End If
        If Not compiler__genCode_(colPolish, _
                                  intIndex1, _
                                  ENUop.opForIncrement, _
                                  Nothing, _
                                  "For loop increment or decrement") Then Return compiler__parseFail_("forNext")
        If Not compiler__genCode_(colPolish, _
                                  intIndex1, _
                                  ENUop.opJump, _
                                  usrNesting.strStartLabel, _
                                  "Jump back to start of For loop") Then
            Return compiler__parseFail_("forNext")
        End If
        Dim strLabel As String = compiler__getLabel_()
        If Not compiler__updateExits_(usrNesting, strLabel, colPolish) Then Return compiler__parseFail_("forNext")
        If Not compiler__genCode_(colPolish, intIndex1, ENUop.opLabel, strLabel, "For loop exit target") Then Return compiler__parseFail_("forNext")
        If Not compiler__genCode_(colPolish, intIndex1, ENUop.opPopOff, Nothing, "Remove the For stack frame") Then Return compiler__parseFail_("forNext")
        If Not compiler__genCode_(colPolish, intIndex1, ENUop.opPopOff, Nothing, "") Then Return compiler__parseFail_("forNext")
        If compiler__genCode_(colPolish, intIndex1, ENUop.opPopOff, Nothing, "") Then
            compiler__parseEvent_("forNext", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("forNext")
    End Function

    ' ----------------------------------------------------------------------
    ' gosub := GoSub (integer | identifier | expression)
    '
    '
    ' Note that because we've tested an integer and an identifier (label)
    ' before testing for an expression, we know that the expression
    ' represents a "computed" gosub and here, we can generate a jump
    ' indirect to the address on the stack.
    '
    '
    Private Function compiler__statementBody__gosub_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "gosub")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__statementBody__gosubGoto_(intIndex, _
                                                   objScanned, _
                                                   colPolish, _
                                                   intEndIndex, _
                                                   strSourceCode, _
                                                   colVariables, _
                                                   "GoSub", _
                                                   intLevel) Then
            compiler__parseEvent_("gosub", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("gosub")
    End Function

    ' ----------------------------------------------------------------------
    ' Common logic for compilation of goto and gosub
    '
    '
    Private Function compiler__statementBody__gosubGoto_(ByRef intIndex As Integer, _
                                                            ByVal objScanned As qbScanner.qbScanner, _
                                                            ByRef colPolish As Collection, _
                                                            ByVal intEndIndex As Integer, _
                                                            ByVal strSourceCode As String, _
                                                            ByRef colVariables As Collection, _
                                                            ByVal strKeyword As String, _
                                                            ByVal intLevel As Integer) As Boolean
        Dim strType As String           ' STMTNUMBER, LABEL or COMPUTED
        Dim strKeywordWork As String = Trim(UCase(strKeyword))
        Dim intIndex1 As Integer = intIndex
        Dim booImplicitGoTo As Boolean
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     strKeywordWork, _
                                     intLevel + 1) Then
            If strKeywordWork <> "GOTO" Then Return (False)
            booImplicitGoTo = True
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger, _
                                 intLevel + 1) Then
            strType = "STMTNUMBER"
        ElseIf booImplicitGoTo Then
            Return (False)
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                     intLevel + 1) Then
            strType = "LABEL"
        Else
            If compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     colVariables, _
                                     intLevel + 1) Then
                strType = "COMPUTED"
            Else
                compiler__errorHandler_(strKeyword & " not followed by gosub target", intIndex, objScanned, strSourceCode)
                Return (False)
            End If
        End If
        With objScanned.QBToken(intIndex - 1)
            Dim strRetLabel As String
            If strKeywordWork = "GOSUB" Then
                strRetLabel = compiler__getLabel_()
                If Not compiler__genCode_(colPolish, intIndex, _
                                            ENUop.opPushReturn, _
                                            strRetLabel, _
                                            "Push GoSub return address") Then Return (False)
            End If
            Dim strLabel As String
            If strType <> "COMPUTED" Then
                strLabel = _OBJutilities.dequote(Mid(strSourceCode, .StartIndex, .Length))
            End If
            If (strType = "COMPUTED" _
                AndAlso _
                (strKeywordWork <> "GOSUB" _
                 OrElse _
                 compiler__genCode_(colPolish, intIndex, _
                                    ENUop.opRotate, _
                                    1, _
                                    "computed tgt,return label->return label,computed tgt")) _
                AndAlso _
                compiler__genCode_(colPolish, intIndex, _
                                   ENUop.opCoGo, _
                                   Nothing, _
                                   "Translate computed " & strKeyword & " target") _
                AndAlso _
                compiler__genCode_(colPolish, intIndex, _
                                   ENUop.opJumpIndirect, _
                                   Nothing, _
                                   "Jump to computed " & strKeyword & " target") _
                OrElse _
                compiler__genCode_(colPolish, intIndex, _
                                   ENUop.opJump, _
                                   strLabel, _
                                   CStr(IIf(strKeywordWork = "GOTO", _
                                            "Go to", _
                                            "Call subroutine at")) & _
                                   " " & _
                                   CStr(IIf(strType = "LABEL", "label", "statement number")) & _
                                   " " & _
                                   strLabel)) _
               AndAlso _
               (strKeywordWork <> "GOSUB" _
                OrElse _
                compiler__genCode_(colPolish, intIndex, _
                                   ENUop.opLabel, _
                                   strRetLabel, _
                                   "Subroutine's return")) Then
                Return (True)
            End If
            Return (False)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' goto := GoTo (integer | identifier | expression)
    ' goto := integer
    '
    '
    ' Note that because we've tested an integer and an identifier (label)
    ' before testing for an expression, we know that the expression
    ' represents a "computed" goto and here, we can generate a jump
    ' indirect to the address on the stack.
    '
    '
    Private Function compiler__statementBody__goto_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "goto")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__statementBody__gosubGoto_(intIndex, _
                                                   objScanned, _
                                                   colPolish, _
                                                   intEndIndex, _
                                                   strSourceCode, _
                                                   colVariables, _
                                                   "GoTo", _
                                                   intLevel) Then
            compiler__parseEvent_("goto", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("goto")
    End Function

    ' ----------------------------------------------------------------------
    ' if := If expression [ Then ] ( unconditionalStatementBody | 
    '                                assignmentStmt |
    '                                functionCall )
    ' if := If expression Then
    '
    '
    Private Function compiler__statementBody__if_(ByRef intIndex As Integer, _
                                                  ByVal objScanned As qbScanner.qbScanner, _
                                                  ByRef colPolish As Collection, _
                                                  ByVal intEndIndex As Integer, _
                                                  ByVal strSourceCode As String, _
                                                  ByRef objNestingStack As Stack, _
                                                  ByRef colVariables As Collection, _
                                                  ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "if")
        Dim intIndexStart As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "IF", _
                                     intLevel + 1) Then Return compiler__parseFail_("if")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer = intIndex
        Dim intStartIndex As Integer = colPolish.Count + 1
        Dim intThenIndex As Integer = intIndex
        For intThenIndex = intThenIndex To intEndIndex
            If objScanned.QBToken(intThenIndex).TokenType _
               = _
               qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline _
               OrElse _
               compiler__checkToken_(intThenIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "THEN", _
                                     intLevel + 1) Then Exit For
        Next intThenIndex
        Dim booSave As Boolean = OBJstate.usrState.booConstantFolding
        OBJstate.usrState.booConstantFolding = False
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intThenIndex - 1, _
                                     strSourceCode, _
                                     colVariables, _
                                     intLevel + 1) Then
            compiler__errorHandler_("If not followed by conditional expression", intIndex, objScanned, strSourceCode)
            Return compiler__parseFail_("if")
        End If
        OBJstate.usrState.booConstantFolding = booSave
        If Not compiler__genCode_(colPolish, _
                                  intIndex, _
                                  ENUop.opJumpZ, _
                                  0, _
                                  "Jump to False code") Then Return compiler__parseFail_("if")
        Dim intFalseJumpIndex As Integer = colPolish.Count
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intThenIndex, _
                                 "THEN", _
                                 intLevel + 1) _
           AndAlso _
           intIndex <= intEndIndex _
           AndAlso _
           objScanned.QBToken(intIndex).TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline Then
            ' Structured If Then
            compiler__pushNesting_(objNestingStack, _
                                   "", _
                                   ENUnesting.ifThen, _
                                   intFalseJumpIndex, _
                                   -1, _
                                   objScanned.QBToken(intIndex).LineNumber, _
                                   "")
        Else
            If compiler__statementBody__unconditional_(intIndex, _
                                                       objScanned, _
                                                       colPolish, _
                                                       intEndIndex, _
                                                       strSourceCode, _
                                                       colVariables, _
                                                       objNestingStack, _
                                                       intLevel + 1) _
                OrElse _
                compiler__statementBody__assignmentStmt_(intIndex, _
                                                         objScanned, _
                                                         colPolish, _
                                                         intEndIndex, _
                                                         strSourceCode, _
                                                         colVariables, _
                                                         intLevel + 1) _
                OrElse _
                compiler__functionCall_(intIndex, _
                                        objScanned, _
                                        colPolish, _
                                        intEndIndex, _
                                        strSourceCode, _
                                        colVariables, _
                                        intLevel + 1) Then
                ' One line If Then
                Dim strLabel As String = compiler__getLabel_()
                If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "End of a one-line If statement") Then
                    Return compiler__parseFail_("if")
                End If
                CType(colPolish(intFalseJumpIndex), qbPolish.qbPolish).Operand = strLabel
            Else
                compiler__errorHandler_("If statement is incomplete", intIndex, objScanned, strSourceCode)
                Return compiler__parseFail_("if")
            End If
        End If
        compiler__parseEvent_("if", _
                                False, _
                                intIndex2, _
                                intIndex - intIndex2, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' input := Input lValueList
    '
    '
    Private Function compiler__statementBody__input_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "input")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "INPUT", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("input")
        End If
        If compiler__lValueList_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 ENUlValueContext.input, _
                                 colVariables, _
                                 intLevel + 1) Then
            compiler__parseEvent_("input", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("input")
    End Function

    ' ----------------------------------------------------------------------
    ' loopOrWend := Wend | ( Loop [ whileUntilClause ] )
    '
    '
    Private Function compiler__statementBody__loopOrWend_(ByRef intIndex As Integer, _
                                                          ByVal objScanned As qbScanner.qbScanner, _
                                                          ByRef colPolish As Collection, _
                                                          ByVal intEndIndex As Integer, _
                                                          ByVal strSourceCode As String, _
                                                          ByRef objNestingStack As Stack, _
                                                          ByRef colVariables As Collection, _
                                                          ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "loopOrWend")
        Dim enuNestType As ENUnesting
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer = intIndex
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "WEND", _
                                 intLevel + 1) Then
            enuNestType = ENUnesting.whileLoop
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "LOOP", _
                                     intLevel + 1) Then
            enuNestType = ENUnesting.doLoop
        Else
            Return compiler__parseFail_("loopOrWend")
        End If
        Dim usrStackNest As TYPnesting
        usrStackNest.enuNestType = ENUnesting.openCode
        Try
            usrStackNest = CType(objNestingStack.Pop, TYPnesting)
        Catch : End Try
        With usrStackNest
            If .enuNestType <> enuNestType Then
                compiler__errorHandler_(CStr(IIf(enuNestType = ENUnesting.doLoop, "Loop", "Wend")) & " " & _
                                        "statement does not balance " & _
                                        compiler__nestingType2String_(.enuNestType), _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
                Return compiler__parseFail_("loopOrWend")
            End If
            If .enuNestType = ENUnesting.doLoop Then
                If Not compiler__whileUntilClause_(intIndex, _
                                                   objScanned, _
                                                   colPolish, _
                                                   intEndIndex, _
                                                   strSourceCode, _
                                                   objNestingStack, _
                                                   colVariables, _
                                                   intLevel + 1) Then
                    If Not compiler__genCode_(colPolish, _
                                              intIndex, _
                                              ENUop.opJump, _
                                              .strStartLabel, _
                                              "Jump to start of unconditional Do loop") Then Return compiler__parseFail_("loopOrWend")
                End If
            Else
                If Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opJump, _
                                          .strStartLabel, _
                                          "Jump to start of while loop") Then Return compiler__parseFail_("loopOrWend")
            End If
            Dim strLabel As String = compiler__getLabel_()
            If Not compiler__updateExits_(usrStackNest, strLabel, colPolish) Then Return compiler__parseFail_("loopOrWend")
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "Loop exit target") Then
                Return compiler__parseFail_("loopOrWend")
            End If
            compiler__parseEvent_("loopOrWend", _
                                    False, _
                                    intIndex2, _
                                    intIndex - intIndex2, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        compiler__parseFail_("loopOrWend")
    End Function

    ' ----------------------------------------------------------------------
    ' print := Print [ expressionList ] [ ; ]
    '
    '
    ' Note: we pass special parameters to expressionList to specify that
    ' the expression ends before the optional semicolons and that the
    ' expressions can be delimited by semicolons and commas.
    '
    '
    Private Function compiler__statementBody__print_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByRef colVariables As Collection, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "print")
        Dim intIndexStart As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "PRINT", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("print")
        End If
        Dim intCountPrevious As Integer
        Dim intIndex1 As Integer = compiler__findNewline_(intIndex, objScanned, intEndIndex, True)
        compiler__expressionList_(intIndex, _
                                  objScanned, _
                                  colPolish, _
                                  intIndex1 - 1, _
                                  strSourceCode, _
                                  colVariables, _
                                  intLevel + 1)
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 ";", _
                                 intLevel + 1) Then
            compiler__parseEvent_("print", _
                                    False, _
                                    intIndexStart, _
                                    intIndex - intIndexStart, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "Print ... ;", _
                                    intLevel)
            Return (True)
        Else
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opPushLiteral, _
                                      _OBJqbVariable.mkVariable _
                                      ("String:" & _OBJutilities.object2String(vbNewLine, True)), _
                                      "Terminate print line") Then Return compiler__parseFail_("print")
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opConcat, _
                                      Nothing, _
                                      "") Then Return compiler__parseFail_("print")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opPrint, Nothing) Then
            compiler__parseEvent_("print", _
                                    False, _
                                    intIndexStart, _
                                    intIndex - intIndexStart, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "Print ...", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("print")
    End Function

    ' ----------------------------------------------------------------------
    ' randomize := Randomize
    '
    '
    Private Function compiler__statementBody__randomize_(ByRef intIndex As Integer, _
                                                         ByVal objScanned As qbScanner.qbScanner, _
                                                         ByRef colPolish As Collection, _
                                                         ByVal intEndIndex As Integer, _
                                                         ByVal strSourceCode As String, _
                                                         ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "randomize")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "RANDOMIZE", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("randomize")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opRand, Nothing) Then
            compiler__parseEvent_("randomize", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("randomize")
    End Function

    ' ----------------------------------------------------------------------
    ' Read := Read lvalueList
    '
    '
    Private Function compiler__statementBody__read_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByRef colVariables As Collection, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "read")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "READ", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("read")
        End If
        If compiler__lValueList_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 ENUlValueContext.read, _
                                 colVariables, _
                                 intLevel + 1) Then
            compiler__parseEvent_("read", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("read")
    End Function

    ' ----------------------------------------------------------------------
    ' Return := Return
    '
    '
    Private Function compiler__statementBody__return_(ByRef intIndex As Integer, _
                                                      ByVal objScanned As qbScanner.qbScanner, _
                                                      ByRef colPolish As Collection, _
                                                      ByVal intEndIndex As Integer, _
                                                      ByVal strSourceCode As String, _
                                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "return")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "RETURN", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("return")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opJumpIndirect, "0", "Return from the subroutine") Then
            compiler__parseEvent_("return", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("return")
    End Function

    ' ----------------------------------------------------------------------
    ' Screen : = SCREEN n
    '
    '
    ' Note that the SCREEN statement generates a clear screen command and
    ' has no other effect.
    '
    '
    Private Function compiler__statementBody__screen_(ByRef intIndex As Integer, _
                                                      ByVal objScanned As qbScanner.qbScanner, _
                                                      ByRef colPolish As Collection, _
                                                      ByVal intEndIndex As Integer, _
                                                      ByVal strSourceCode As String, _
                                                      ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "screen")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "SCREEN", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("screen")
        End If
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger, _
                                     intLevel + 1) Then
            intIndex = intIndex1
            compiler__errorHandler_("SCREEN is not followed by an integer code", _
                                    intIndex, _
                                    objScanned, _
                                    strSourceCode)
            Return compiler__parseFail_("screen")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opCls, "0", "SCREEN command") Then
            compiler__parseEvent_("screen", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("screen")
    End Function

    ' ----------------------------------------------------------------------
    ' stop := Stop
    '
    '
    Private Function compiler__statementBody__stop_(ByRef intIndex As Integer, _
                                                    ByVal objScanned As qbScanner.qbScanner, _
                                                    ByRef colPolish As Collection, _
                                                    ByVal intEndIndex As Integer, _
                                                    ByVal strSourceCode As String, _
                                                    ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "stop")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "STOP", _
                                     intLevel + 1) Then
            Return compiler__parseFail_("stop")
        End If
        If compiler__genCode_(colPolish, intIndex, ENUop.opEnd, Nothing) Then
            compiler__parseEvent_("stop", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        Return compiler__parseFail_("stop")
    End Function

    ' ----------------------------------------------------------------------
    ' trace := Trace Push 
    ' trace := Trace Off 
    ' trace := Trace Text (Source|Memory|Stack|Inst|Object|Line|n|NoBox)*    
    ' trace := Trace Headsup (Inst|Line|n)*   
    ' trace := Trace HeadsupText (Source|Memory|Stack|Inst|Object|Line|n|NoBox)*    
    ' trace := Trace Pop 
    '
    '
    Private Function compiler__statementBody__trace_(ByRef intIndex As Integer, _
                                                     ByVal objScanned As qbScanner.qbScanner, _
                                                     ByRef colPolish As Collection, _
                                                     ByVal intEndIndex As Integer, _
                                                     ByVal strSourceCode As String, _
                                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "trace")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "TRACE", _
                                     intLevel + 1) Then Return compiler__parseFail_("trace")
        Dim strOperands As String
        If intIndex < intEndIndex Then
            strOperands = rebuildCode_(objScanned, _
                                        strSourceCode, _
                                        intIndex + 1, _
                                        compiler__findNewline_(intIndex, _
                                                            objScanned, _
                                                            intEndIndex, _
                                                            False) _
                                        - _
                                        1)
        End If
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "PUSH", _
                                 intLevel + 1) Then
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opTracePush, _
                                      0, _
                                      "Stack the trace settings") Then Return compiler__parseFail_("trace")
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "TEXT", _
                                     intLevel + 1) Then
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opTrace, _
                                      traceStruct2Code_(compiler__statementBody__trace__parser_(strOperands, _
                                                                                                True, False, _
                                                                                                intIndex, _
                                                                                                objScanned, _
                                                                                                strSourceCode), _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode), _
                                      "Set text tracing on") Then Return compiler__parseFail_("trace")
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "HEADSUP", _
                                     intLevel + 1) Then
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opTrace, _
                                      traceStruct2Code_(compiler__statementBody__trace__parser_(strOperands, _
                                                                                                False, True, _
                                                                                                intIndex, _
                                                                                                objScanned, _
                                                                                                strSourceCode), _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode), _
                                      "Set headsup tracing on") Then Return compiler__parseFail_("trace")
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "HEADSUPTEXT", _
                                     intLevel + 1) Then
            If Not compiler__genCode_(colPolish, _
                                      intIndex, _
                                      ENUop.opTrace, _
                                      traceStruct2Code_(compiler__statementBody__trace__parser_(strOperands, _
                                                                                                True, True, _
                                                                                                intIndex, _
                                                                                                objScanned, _
                                                                                                strSourceCode), _
                                                        intIndex, _
                                                        objScanned, _
                                                        strSourceCode), _
                                      "Set headsup tracing on") Then Return compiler__parseFail_("trace")
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "OFF", _
                                     intLevel + 1) Then
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opTrace, 0, "Clear tracing") Then Return compiler__parseFail_("trace")
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "POP", _
                                     intLevel + 1) Then
            If Not compiler__genCode_(colPolish, intIndex, ENUop.opTracePop, 0, "Restore previous trace setting") Then Return compiler__parseFail_("trace")
        Else
            compiler__errorHandler_("Invalid Trace statement", intIndex, objScanned, strSourceCode)
        End If
        compiler__parseEvent_("trace", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Parse the Trace statement operand
    '
    '
    Private Function compiler__statementBody__trace__parser_(ByVal strOperands As String, _
                                                             ByVal booTextTrace As Boolean, _
                                                             ByVal booHeadsUpTrace As Boolean, _
                                                             ByRef intIndex As Integer, _
                                                             ByVal objScanned As qbScanner.qbScanner, _
                                                             ByVal strSourceCode As String) As TYPtracing
        Dim intIndex1 As Integer
        Dim strNext As String
        Dim usrTracing As TYPtracing
        With usrTracing
            .booString2Box = True : .booTextTrace = booTextTrace : .booHeadsupTrace = booHeadsUpTrace
            .intRate = 1
            For intIndex1 = 1 To _OBJutilities.words(strOperands)
                strNext = UCase(_OBJutilities.word(strOperands, intIndex1))
                If booTextTrace Then
                    Select Case strNext
                        Case "SOURCE" : .booIncludeSource = True
                        Case "MEMORY" : .booIncludeMemory = True
                        Case "STACK" : .booIncludeStack = True
                        Case "NOBOX" : .booString2Box = False
                        Case "OBJECT" : .booIncludeObject = True
                        Case Else
                            Select Case strNext
                                Case "INST"
                                Case "LINES" : .booTraceLines = True
                                Case Else
                                    .intRate = -1
                                    If _OBJutilities.verify(strNext, "0123456789") = 0 Then
                                        Try
                                            .intRate = CInt(strNext)
                                        Catch : End Try
                                    End If
                                    If .intRate < 0 Then
                                        compiler__errorHandler_("In the Trace statement, the token " & _
                                                                _OBJutilities.enquote(strNext) & " " & _
                                                                "is not recognizable", _
                                                                intIndex, _
                                                                objScanned, _
                                                                strSourceCode)
                                        .intRate = 1
                                        Return (usrTracing)
                                    End If
                            End Select
                    End Select
                End If
            Next intIndex1
            intIndex += intIndex1 - 1
            Return (usrTracing)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' unconditionalStatementBody := circle | comment | data | end | exit | 
    '                               gosub | goto | input | print | randomize |
    '                               read | return | screen | stop | trace
    '
    '
    Private Function compiler__statementBody__unconditional_(ByRef intIndex As Integer, _
                                                             ByVal objScanned As qbScanner.qbScanner, _
                                                             ByRef colPolish As Collection, _
                                                             ByVal intEndIndex As Integer, _
                                                             ByVal strSourceCode As String, _
                                                             ByRef colVariables As Collection, _
                                                             ByRef objNestingStack As Stack, _
                                                             ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "unconditional")
        Dim booReturn As Boolean
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If compiler__statementBody__circle_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__cls_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__comment_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__data_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__end_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__exit_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, objNestingStack, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__gosub_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__goto_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__input_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__print_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__randomize_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__read_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, colVariables, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__return_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__stop_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__screen_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        ElseIf compiler__statementBody__trace_(intIndex, objScanned, colPolish, intEndIndex, strSourceCode, intLevel + 1) Then
            booReturn = True
        End If
        If booReturn Then
            compiler__parseEvent_("unconditional", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End If
        compiler__parseFail_("unconditional")
    End Function

    ' ----------------------------------------------------------------------
    ' whileHeader := While expression 
    '
    '
    Private Function compiler__statementBody__whileHeader_(ByRef intIndex As Integer, _
                                                           ByVal objScanned As qbScanner.qbScanner, _
                                                           ByRef colPolish As Collection, _
                                                           ByVal intEndIndex As Integer, _
                                                           ByVal strSourceCode As String, _
                                                           ByRef objNestingStack As Stack, _
                                                           ByRef colVariables As Collection, _
                                                           ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "whileHeader")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex1 As Integer = intIndex
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "WHILE", _
                                     intLevel + 1) Then Return compiler__parseFail_("whileHeader")
        Dim strLabel As String = compiler__getLabel_()
        If Not compiler__genCode_(colPolish, intIndex, ENUop.opLabel, strLabel, "") Then Return compiler__parseFail_("whileHeader")
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     colVariables, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("whileHeader")
        End If
        If Not compiler__pushNesting_(objNestingStack, _
                                      strLabel, _
                                      ENUnesting.whileLoop, _
                                      Nothing, _
                                      0, _
                                      objScanned.QBToken(intIndex).LineNumber, _
                                      "") Then
            Return compiler__parseFail_("trace")
        End If
        compiler__parseEvent_("whileHeader", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert and check object subscript to integer, on behalf of
    ' compiler__subscriptList: return -1 on error
    '
    '
    Private Function compiler__expressionList_object2IntegerSub_(ByVal objSubscript As qbVariable.qbVariable, _
                                                                 ByVal booDimCompile As Boolean, _
                                                                 ByVal intIndex As Integer, _
                                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                                 ByVal strSourceCode As String) As Integer
        Dim intSubscript As Integer
        Dim strSubscriptType As String = CStr(IIf(booDimCompile, "upper bound", "subscript"))
        Try
            intSubscript = CInt(objSubscript.value)
        Catch
            compiler__errorHandler_("Unusable " & strSubscriptType & " " & _
                                    _OBJutilities.object2String(objSubscript) & " " & _
                                    "specified for array", _
                                    intIndex, objScanned, strSourceCode)
            Return (-1)
        End Try
        If intSubscript < 0 Then
            compiler__errorHandler_("Invalid " & strSubscriptType & " " & _
                                    intSubscript & " " & _
                                    "specified for array", _
                                    intIndex, objScanned, strSourceCode)
            Return (-1)
        End If
        Return (intSubscript)
    End Function

    ' ----------------------------------------------------------------------
    ' term := number | string | lValue | True | False | functionCall | ( expression )
    '
    '
    Private Function compiler__term_(ByRef intIndex As Integer, _
                                     ByVal objScanned As qbScanner.qbScanner, _
                                     ByRef colPolish As Collection, _
                                     ByVal intEndIndex As Integer, _
                                     ByVal strSourceCode As String, _
                                     ByRef objConstantValue As qbVariable.qbVariable, _
                                     ByRef colVariables As Collection, _
                                     ByRef booSideEffects As Boolean, _
                                     ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "term")
        Dim dblValue As Double
        Dim objType As qbVariableType.qbVariableType
        Dim intAddress As Integer
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndexStart As Integer = intIndex
        If intIndex > intEndIndex Then Return compiler__parseFail_("term")
        If compiler__number_(intIndex, _
                             objScanned, _
                             intEndIndex, _
                             strSourceCode, dblValue, _
                             objType, _
                             intLevel + 1) Then
            If OBJstate.usrState.booConstantFolding Then
                If Not compiler__term__mkConstantValue_(dblValue, objConstantValue) Then
                    Return compiler__parseFail_("term")
                End If
            Else
                If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, dblValue, "Push numeric constant") Then Return compiler__parseFail_("term")
            End If
        ElseIf objScanned.QBToken(intIndex).TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeString Then
            Dim strValue As String = Mid(strSourceCode, _
                                         objScanned.QBToken(intIndex).StartIndex + 1, _
                                         objScanned.QBToken(intIndex).Length - 2)
            If OBJstate.usrState.booConstantFolding Then
                If Not compiler__term__mkConstantValue_(strValue, objConstantValue) Then
                    Return compiler__parseFail_("term")
                End If
            Else
                If Not compiler__genCode_(colPolish, intIndex, ENUop.opPushLiteral, strValue, "Push string constant") Then
                    Return compiler__parseFail_("term")
                End If
            End If
            intIndex += 1
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "TRUE", _
                                     intLevel + 1) Then
            If OBJstate.usrState.booConstantFolding Then
                If Not compiler__term__mkConstantValue_(True, objConstantValue) Then
                    Return compiler__parseFail_("term")
                End If
            ElseIf Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opPushLiteral, _
                                          -1, _
                                          "Push the True") Then
                Return compiler__parseFail_("term")
            End If
        ElseIf compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     "FALSE", _
                                     intLevel + 1) Then
            If OBJstate.usrState.booConstantFolding Then
                If Not compiler__term__mkConstantValue_(False, objConstantValue) Then
                    Return compiler__parseFail_("term")
                End If
            ElseIf Not compiler__genCode_(colPolish, _
                                          intIndex, _
                                          ENUop.opPushLiteral, _
                                          0, _
                                          "Push the False") Then
                Return compiler__parseFail_("term")
            End If
        ElseIf compiler__functionCall_(intIndex, _
                                       objScanned, _
                                       colPolish, _
                                       intEndIndex, _
                                       strSourceCode, _
                                       colVariables, _
                                       booSideEffects, _
                                       intLevel + 1) Then
            objConstantValue = Nothing
        Else
            Dim intIndex1 As Integer = intIndex
            If compiler__lValue_(intIndex, _
                                 objScanned, _
                                 colPolish, _
                                 intEndIndex, _
                                 strSourceCode, _
                                 intAddress, _
                                 colVariables, _
                                 False, _
                                 intLevel + 1) Then
                If Not compiler__genCode_(colPolish, intIndex, _
                                          ENUop.opNop, _
                                          0, _
                                          "Push lValue " & _
                                          rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1) & " " & _
                                          CStr(IIf(intAddress = -1, "indirect", "")) & " " & _
                                          "contents of memory location") Then Return compiler__parseFail_("term")
                If intAddress = MEMORY_LOCATION_STACKEDADDRESS Then
                    ' Indirect reference at top of stack
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              ENUop.opPushIndirect, _
                                              Nothing, _
                                              "Push contents of memory location") Then Return compiler__parseFail_("term")
                ElseIf intAddress = MEMORY_LOCATION_STACKFRAME Then
                    ' Array frame
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              ENUop.opPushArrayElement, _
                                              0, _
                                              "Push array element") Then Return compiler__parseFail_("term")
                Else
                    ' Have address
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              ENUop.opPushLiteral, _
                                              intAddress, _
                                              "Push indirect address") Then Return compiler__parseFail_("term")
                    If Not compiler__genCode_(colPolish, intIndex, _
                                              ENUop.opPushIndirect, _
                                              Nothing, _
                                              "Push contents of memory location") Then Return compiler__parseFail_("term")
                End If
            ElseIf objScanned.QBToken(intIndex).TokenType _
                   = _
                   qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis _
                AndAlso _
                Mid(strSourceCode, _
                    objScanned.QBToken(intIndex).StartIndex, _
                    objScanned.QBToken(intIndex).Length) _
                = _
                "(" Then
                intIndex += 1
                If compiler__expression_(intIndex, _
                                            objScanned, _
                                            colPolish, _
                                            objScanned.findRightParenthesis(intIndex, _
                                                                            intEndIndex:=intEndIndex) _
                                            - _
                                            1, _
                                            strSourceCode, _
                                            objConstantValue, _
                                            colVariables, _
                                            intLevel + 1) Then
                    intIndex += 1
                End If
            Else
                compiler__errorHandler_("Unexpected string cannot be recognized by the parser", _
                                        intIndex, objScanned, strSourceCode)
                Return compiler__parseFail_("term")
            End If
        End If
        compiler__parseEvent_("term", _
                                False, _
                                intIndexStart, _
                                intIndex - intIndexStart, _
                                intCountPrevious + 1, _
                                colPolish.Count - intCountPrevious, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Make a constant value on behalf of compiler__term_
    '
    '
    Private Function compiler__term__mkConstantValue_(ByVal objValue As Object, _
                                                      ByRef objConstantValue _
                                                            As qbVariable.qbVariable) _
            As Boolean
        Try
            If (Not objConstantValue Is Nothing) Then
                disposeVariable_(objConstantValue)
                objConstantValue = Nothing
            End If
            objConstantValue = _OBJqbVariable.mkVariableFromValue(CStr(objValue))
        Catch objException As Exception
            errorHandler_("Cannot make a constant value as a qbVariable: " & _
                          Err.Number & " " & Err.Description, _
                          "", _
                          "Marking object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False : Return (False)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Operator token to opcode as used by the interpreter
    '
    '
    Private Function compiler__token2Op_(ByVal strToken As String) As qbOp.qbOp.ENUop
        Select Case UCase(Trim(strToken))
            Case "&" : Return (ENUop.opConcat)
            Case "+" : Return (ENUop.opAdd)
            Case "-" : Return (ENUop.opSubtract)
            Case "*" : Return (ENUop.opMultiply)
            Case "/" : Return (ENUop.opDivide)
            Case "\" : Return (ENUop.opDivide)
            Case "^" : Return (ENUop.opPower)
            Case "**" : Return (ENUop.opPower)
            Case "=" : Return (ENUop.opPushEQ)
            Case "<>" : Return (ENUop.opPushNE)
            Case "<" : Return (ENUop.opPushLT)
            Case "<=" : Return (ENUop.opPushLE)
            Case ">" : Return (ENUop.opPushGT)
            Case ">=" : Return (ENUop.opPushGE)
            Case "OR" : Return (ENUop.opOr)
            Case "AND" : Return (ENUop.opAnd)
            Case "MOD" : Return (ENUop.opMod)
            Case "LIKE" : Return (ENUop.opLike)
            Case Else
                errorHandler_("Internal compiler error: unrecognizable operator", _
                              "compiler__token2Op_", _
                              "Returning NOP and making object unusable", _
                              Nothing)
                OBJstate.usrState.booUsable = False
                Return (ENUop.opNop)
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Convert type character to variable type enumerator
    '
    '
    Private Function compiler__typechar2Type_(ByVal intIndex As Integer, _
                                              ByVal objScanned As qbScanner.qbScanner, _
                                              ByVal strSource As String, _
                                              ByVal strType As String) As String
        Select Case strType
            Case POSTFIX_DATATYPE_CHAR_SHORT
                Return (qbVariableType.qbVariableType.ENUvarType.vtInteger.ToString)
            Case POSTFIX_DATATYPE_CHAR_LONG
                Return (qbVariableType.qbVariableType.ENUvarType.vtLong.ToString)
            Case POSTFIX_DATATYPE_CHAR_SINGLE
                Return (qbVariableType.qbVariableType.ENUvarType.vtSingle.ToString)
            Case POSTFIX_DATATYPE_CHAR_DOUBLE
                Return (qbVariableType.qbVariableType.ENUvarType.vtDouble.ToString)
            Case POSTFIX_DATATYPE_CHAR_STRING
                Return (qbVariableType.qbVariableType.ENUvarType.vtString.ToString)
            Case Else
                compiler__errorHandler_("Invalid type character " & _
                                        _OBJutilities.enquote(strType), _
                                        intIndex, _
                                        objScanned, _
                                        strSource, _
                                        "The valid type characters are " & _
                                        POSTFIX_DATATYPE_CHARS)
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' typedIdentifier := identifier [ typeCharacter ]
    '
    '
    Private Function compiler__typedIdentifier_(ByRef intIndex As Integer, _
                                                ByVal objScanned As qbScanner.qbScanner, _
                                                ByVal intEndIndex As Integer, _
                                                ByVal strSourceCode As String, _
                                                ByRef strIdentifier As String, _
                                                ByRef objType As qbVariableType.qbVariableType, _
                                                ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "typedIdentifier")
        Dim intIndex1 As Integer = intIndex
        strIdentifier = "" : objType.fromString("Variant")
        If Not compiler__checkToken_(intIndex, _
                                     objScanned, _
                                     strSourceCode, _
                                     intEndIndex, _
                                     qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("typedIdentifier")
        End If
        With objScanned.QBToken(intIndex - 1)
            strIdentifier = Mid(strSourceCode, .StartIndex, .Length)
        End With
        compiler__typeSuffix_(intIndex, _
                              objScanned, _
                              intEndIndex, _
                              strSourceCode, _
                              objType, _
                              intLevel + 1)
        compiler__parseEvent_("typedIdentifier", _
                                False, _
                                intIndex1, _
                                intIndex - intIndex1, _
                                0, 0, _
                                "", _
                                intLevel)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' typeSuffix := numTypeChar | currencySymbol
    '
    '
    ' Note: the type suffix must appear after the first token, and it must be
    ' adjacent to the preceding token, which is assumed by this method to be
    ' an identifier.
    ' 
    Private Function compiler__typeSuffix_(ByRef intIndex As Integer, _
                                           ByVal objScanned As qbScanner.qbScanner, _
                                           ByVal intEndIndex As Integer, _
                                           ByVal strSourceCode As String, _
                                           ByRef objType As qbVariableType.qbVariableType, _
                                           ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "typeSuffix")
        If intIndex < 2 Then Return compiler__parseFail_("typeSuffix")
        Dim intIndex1 As Integer = intIndex
        Dim intIndex2 As Integer = intIndex
        objType.fromString("Variant")
        If compiler__numTypeChar_(intIndex2, _
                                  objScanned, _
                                  intEndIndex, _
                                  strSourceCode, _
                                  objType, _
                                  intLevel + 1) _
           AndAlso _
           objScanned.tokenStartIndex(intIndex2 - 1) = objScanned.tokenEndIndex(intIndex2 - 2) Then
            compiler__parseEvent_("typeSuffix", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    0, 0, _
                                    "Numeric type character", _
                                    intLevel)
            intIndex = intIndex2
            Return (True)
        End If
        intIndex2 = intIndex1
        If compiler__checkToken_(intIndex2, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeCurrency, _
                                 intLevel + 1) _
           AndAlso _
           objScanned.tokenStartIndex(intIndex2 - 1) = objScanned.tokenEndIndex(intIndex2 - 2) Then
            With objScanned.QBToken(intIndex - 1)
                objType.fromString(compiler__typechar2Type_(intIndex, _
                                                            objScanned, _
                                                            strSourceCode, _
                                                            Mid(strSourceCode, .StartIndex, .Length)))
            End With
            compiler__parseEvent_("typeSuffix", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    0, 0, _
                                    "Currency symbol", _
                                    intLevel)
            intIndex = intIndex2
            Return (True)
        End If
        Return compiler__parseFail_("typeSuffix")
    End Function

    ' ----------------------------------------------------------------------
    ' Update loop exits
    '
    '
    Private Function compiler__updateExits_(ByVal usrNesting As TYPnesting, _
                                            ByVal strLabel As String, _
                                            ByRef colPolish As Collection) As Boolean
        Dim intIndex1 As Integer
        For intIndex1 = 1 To UBound(usrNesting.intLoopExits)
            CType(colPolish.Item(usrNesting.intLoopExits(intIndex1)), _
                  qbPolish.qbPolish).Operand = _
                  strLabel
        Next intIndex1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' whileUntilClause := While|Until expression
    '
    '
    Private Function compiler__whileUntilClause_(ByRef intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByRef colPolish As Collection, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByVal strSourceCode As String, _
                                                 ByRef objNestingStack As Stack, _
                                                 ByRef colVariables As Collection, _
                                                 ByVal intLevel As Integer) As Boolean
        raiseEvent_("parseStartEvent", "whileUntilClause")
        Dim intCountPrevious As Integer = colPolish.Count
        Dim intIndex2 As Integer = intIndex
        Dim booWhile As Boolean
        If compiler__checkToken_(intIndex, _
                                 objScanned, _
                                 strSourceCode, _
                                 intEndIndex, _
                                 "WHILE", _
                                 intLevel + 1) Then
            booWhile = True
        ElseIf Not compiler__checkToken_(intIndex, _
                                         objScanned, _
                                         strSourceCode, _
                                         intEndIndex, _
                                         "UNTIL", _
                                         intLevel + 1) Then
            Return compiler__parseFail_("whileUntilClause")
        End If
        Dim intIndex1 As Integer = intIndex
        If Not compiler__expression_(intIndex, _
                                     objScanned, _
                                     colPolish, _
                                     intEndIndex, _
                                     strSourceCode, _
                                     colVariables, _
                                     intLevel + 1) Then
            Return compiler__parseFail_("whileUntilClause")
        End If
        Dim usrNesting As TYPnesting = CType(objNestingStack.Peek, TYPnesting)
        With usrNesting
            Dim booZeroTrip As Boolean = (.enuNestType <> ENUnesting.doLoop)
            Dim enuOpCode As qbOp.qbOp.ENUop
            Dim strOperand As String
            Dim strWhileUntil As String
            If booWhile AndAlso booZeroTrip Then
                enuOpCode = ENUop.opJumpZ : strOperand = "" : strWhileUntil = "While"
            ElseIf booWhile AndAlso Not booZeroTrip Then
                enuOpCode = ENUop.opJumpNZ : strOperand = .strStartLabel : strWhileUntil = "While"
            ElseIf Not booWhile AndAlso booZeroTrip Then
                enuOpCode = ENUop.opJumpZ : strOperand = "" : strWhileUntil = "Until"
            Else
                enuOpCode = ENUop.opJumpNZ : strOperand = .strStartLabel : strWhileUntil = "Until"
            End If
            If Not compiler__genCode_(colPolish, intIndex, _
                                        enuOpCode, _
                                        CObj(strOperand), _
                                        "Do " & strWhileUntil & " " & _
                                        rebuildCode_(objScanned, strSourceCode, intIndex1, intIndex - 1)) Then
                Return compiler__parseFail_("whileUntilClause")
            End If
            compiler__parseEvent_("term", _
                                    False, _
                                    intIndex1, _
                                    intIndex - intIndex1, _
                                    intCountPrevious + 1, _
                                    colPolish.Count - intCountPrevious, _
                                    "", _
                                    intLevel)
            Return (True)
        End With
        Return compiler__parseFail_("whileUntilClause")
    End Function

#End Region

    ' -----------------------------------------------------------------
    ' Dereference variable index
    '
    '
    ' If objIndex is an integer, this method checks to make sure it is
    ' within the variable collection. If objIndex is a string this method
    ' makes sure that it is the key of a variable.
    '
    ' This method returns the numeric variable index.
    '
    '
    Private Function derefVariableIndex_(ByVal objIndex As Object) As Integer
        With OBJstate.usrState
            Dim intIndex1 As Integer
            If (TypeOf objIndex Is Integer) Then
                intIndex1 = CInt(objIndex)
                If intIndex1 < 1 OrElse intIndex1 > .colVariables.Count Then
                    errorHandler_("Invalid variable index " & intIndex1, _
                                    "derefVariableIndex_", _
                                    "Returning 0", _
                                    Nothing)
                    Return 0
                End If
            ElseIf (TypeOf objIndex Is String) Then
                Dim strIndex As String = CStr(objIndex)
                Try
                    Dim objVariable As qbVariable.qbVariable = _
                        CType(.colVariables.Item(strIndex), qbVariable.qbVariable)
                    intIndex1 = CInt(objVariable.Tag)
                Catch
                    errorHandler_("Invalid variable name " & _
                                    _OBJutilities.enquote(strIndex), _
                                    "derefVariableIndex_", _
                                    "Returning 0", _
                                    Nothing)
                    Return 0
                End Try
            Else
                errorHandler_("Invalid variable index " & _
                                _OBJutilities.object2String(objIndex), _
                                "derefVariableIndex_", _
                                "Returning 0", _
                                Nothing)
                Return 0
            End If
            Return intIndex1
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Dispatcher: includes the threading logic for all cases
    '
    '
    ' --- Don't return a reference value: note that booFlag is only a placeholder
    Private Overloads Function dispatch_(ByVal strProcedure As String, _
                                            ByVal booFlag As Boolean, _
                                            ByVal objDefault As Object, _
                                            ByVal strDefaultHelp As String, _
                                            ByVal ParamArray objParameters() As Object) _
            As Object
        Dim strDummy As String
        Select Case UBound(objParameters)
            Case -1
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp))
            Case 0
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp, _
                                   objParameters(0)))
            Case 1
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp, _
                                   objParameters(0), _
                                   objParameters(1)))
            Case 2
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp, _
                                   objParameters(0), _
                                   objParameters(1), _
                                   objParameters(2)))
            Case 3
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp, _
                                   objParameters(0), _
                                   objParameters(1), _
                                   objParameters(2), _
                                   objParameters(3)))
            Case 4
                Return (dispatch_(strProcedure, _
                                   strDummy, _
                                   objDefault, _
                                   strDefaultHelp, _
                                   objParameters(0), _
                                   objParameters(1), _
                                   objParameters(2), _
                                   objParameters(3), _
                                   objParameters(4)))
            Case Else
                errorHandler_("Internal programming error: " & _
                              "too many parameters", _
                              "dispatch_", _
                              "Making object unusable and returning Nothing", _
                              Nothing)
                OBJstate.usrState.booUsable = False
                Return (Nothing)
        End Select
    End Function
    ' --- Returns the reference value
    Private Overloads Function dispatch_(ByVal strProcedure As String, _
                                            ByRef strOutstring As String, _
                                            ByVal objDefault As Object, _
                                            ByVal strDefaultHelp As String, _
                                            ByVal ParamArray objParameters() As Object) _
            As Object
        SyncLock OBJthreadStatus
            If Not checkAvailability_(strProcedure, strDefaultHelp) Then
                Return (objDefault)
            End If
            OBJthreadStatus.startThread()
        End SyncLock
        Dim objReturn As Object = objDefault
        SyncLock OBJstate
            If checkUsable_(strProcedure, strDefaultHelp) Then
                With OBJstate.usrState
                    Select Case UCase(strProcedure)
                        Case "ASSEMBLE"
                            objReturn = assemble_()
                        Case "ASSEMBLED"
                            objReturn = .booAssembled
                        Case "ASSEMBLYREMOVESCODE GET"
                            objReturn = .booAssemblyRemovesCode
                        Case "ASSEMBLYREMOVESCODE SET"
                            .booAssemblyRemovesCode = CBool(objParameters(0))
                        Case "CLEAR"
                            objReturn = clear_()
                        Case "CLEARSTORAGE"
                            objReturn = clearStorage_()
                        Case "CLONE"
                            objReturn = clone_()
                        Case "COMPARETO"
                            objReturn = compareTo_(CType(objParameters(0), qbQuickBasicEngine))
                        Case "COMPILE"
                            objReturn = compile_()
                        Case "COMPILED"
                            objReturn = .booCompiled
                        Case "CONSTANTFOLDING GET"
                            objReturn = .booConstantFolding
                        Case "CONSTANTFOLDING SET"
                            .booConstantFolding = CBool(objParameters(0))
                        Case "DEGENERATEOPREMOVAL GET"
                            objReturn = .booDegenerateOpRemoval
                        Case "DEGENERATEOPREMOVAL SET"
                            .booDegenerateOpRemoval = CBool(objParameters(0))
                        Case "DISPOSE"
                            objReturn = dispose_(CBool(objParameters(0)))
                        Case "EVALUATE"
                            objReturn = evaluate_(CStr(objParameters(0)), _
                                                  CBool(objParameters(1)), _
                                                  strOutstring)
                        Case "EVALUATION GET"
                            objReturn = .objImmediateResult
                        Case "EVALUATIONVALUE"
                            objReturn = .objImmediateResult.value
                        Case "EVENTLOG GET"
                            objReturn = .colEventLog
                        Case "EVENTLOG2ERRORLIST"
                            objReturn = CStr(eventLog2ErrorList_(.colEventLog))
                        Case "EVENTLOGFORMAT"
                            objReturn = CStr(eventLogFormat_(.colEventLog, _
                                                             CInt(objParameters(0)), _
                                                             CInt(objParameters(1))))
                        Case "EVENTLOGGING GET"
                            objReturn = eventLoggingGet_()
                        Case "EVENTLOGGING SET"
                            eventLoggingSet_(CBool(objParameters(0)))
                        Case "GENERATENOPS GET"
                            objReturn = generateNOPsGet_()
                        Case "GENERATENOPS SET"
                            generateNOPsSet_(CBool(objParameters(0)))
                        Case "INSPECT"
                            objReturn = inspect_(strOutstring)
                        Case "INSPECTCOMPILEROBJECTS GET"
                            objReturn = .booInspectCompilerObjects
                        Case "INSPECTCOMPILEROBJECTS SET"
                            .booInspectCompilerObjects = CBool(objParameters(0))
                        Case "INTERPRET"
                            objReturn = interpret_()
                        Case "MKUNUSABLE"
                            .booUsable = False : objReturn = True
                        Case "MSILRUN"
                            objReturn = msilRun_()
                        Case "NAME GET"
                            objReturn = .strName
                        Case "NAME SET"
                            .strName = CStr(objParameters(0))
                        Case "OBJECT2XML"
                            objReturn = object2XML_(CBool(objParameters(0)), _
                                                    CBool(objParameters(0)))
                        Case "POLISHCOLLECTION GET"
                            objReturn = .colPolish
                        Case "QBSCANNER GET"
                            objReturn = .objScanner
                        Case "QBSCANNER SET"
                            .objScanner = CType(objParameters(0), qbScanner.qbScanner)
                        Case "RESET"
                            objReturn = reset_()
                        Case "RUN"
                            objReturn = run_(CStr(objParameters(0)))
                        Case "SCAN"
                            objReturn = .objScanner.scan
                        Case "SCANNED"
                            objReturn = Not (.objScanner Is Nothing) _
                                        AndAlso _
                                        .objScanner.Scanned
                        Case "SCANNER GET"
                            objReturn = .objScanner
                        Case "SOURCECODE GET"
                            objReturn = ""
                            If Not (.objScanner Is Nothing) Then
                                objReturn = .objScanner.SourceCode
                            End If
                        Case "SOURCECODE GET"
                            objReturn = ""
                            If Not (.objScanner Is Nothing) Then
                                objReturn = .objScanner.SourceCode
                            End If
                        Case "SOURCECODE SET"
                            sourceCodeSet_(CStr(objParameters(0)))
                        Case "TAG GET"
                            objReturn = .objTag
                        Case "TAG SET"
                            .objTag = objParameters(0)
                        Case "TEST"
                            objReturn = test_(strOutstring, CBool(objParameters(0)))
                        Case "VARIABLE GET"
                            objReturn = variableGet_(objParameters(0))
                        Case "VARIABLE SET"
                            objReturn = variableSet_(objParameters(0), _
                                                     CType(objParameters(1), qbVariable.qbVariable))
                        Case "VARIABLECOUNT GET"
                            objReturn = .colVariables.Count
                        Case Else
                            errorHandler_("Invalid dispatch method " & _
                                          _OBJutilities.enquote(strProcedure), _
                                          "dispatch_", _
                                          "Marking object unusable and returning default", _
                                          Nothing)
                            .booUsable = False
                    End Select
                End With
            End If
        End SyncLock
        SyncLock OBJthreadStatus
            OBJthreadStatus.stopThread()
        End SyncLock
        Return (objReturn)
    End Function

    ' -----------------------------------------------------------------
    ' Dispose of the engine safely
    '
    '
    Private Function dispose_(ByVal booInspect As Boolean) As Boolean
        With OBJstate.usrState
            If .booInspectCompilerObjects Then inspection_()
            If Not (.colEventLog Is Nothing) _
            AndAlso _
            Not .objCollectionUtilities.collectionClear(.colEventLog) Then
                Return (False)
            End If
            If Not clear_() Then Return (False)
            .booUsable = False
            If Not .objCollectionUtilities.dispose Then Return (False)
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Dispose of the qbVariable
    '
    '
    Private Function disposeVariable_(ByVal objVariable As qbVariable.qbVariable) _
            As Boolean
        If OBJstate.usrState.booInspectCompilerObjects Then
            Return objVariable.disposeInspect
        Else
            objVariable.dispose()
            Return True
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Interfaces to the error handler
    '
    '
    ' --- Unshared with exception information
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String, _
                                        ByVal objException As Exception)
        errorHandler__(strMessage, OBJstate.usrState.strName, strProcedure, strHelp, objException)
    End Sub
    ' --- Shared
    Private Overloads Shared Sub errorHandler_(ByVal strMessage As String, _
                                               ByVal strProcedure As String, _
                                               ByVal strHelp As String)
        errorHandler__(strMessage, ClassName, strProcedure, strHelp, Nothing)
    End Sub
    ' --- Common logic
    Private Shared Sub errorHandler__(ByVal strMessage As String, _
                                        ByVal strClassOrObjectName As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String, _
                                        ByVal objException As Exception)
        Dim strException As String
        If Not (objException Is Nothing) Then
            strException = vbNewLine & vbNewLine & objException.ToString
        End If
        _OBJutilities.errorHandler(strMessage, _
                                   strClassOrObjectName, _
                                   strProcedure, _
                                   strHelp & _
                                   strException)
    End Sub

    ' ----------------------------------------------------------------------
    ' Make the compiler/interpreter error log from an event log
    '
    '
    Private Shared Function eventLog2ErrorList_(ByVal colEventLog As Collection) _
            As String
        With colEventLog
            Dim colHandle As Collection
            Dim intIndex1 As Integer
            Dim strNext As String
            Dim strNextUCase As String
            Dim strOutstring As String = ""
            For intIndex1 = 1 To .Count
                colHandle = CType(.Item(intIndex1), Collection)
                strNext = CStr(colHandle.Item(1))
                strNextUCase = UCase(strNext)
                If strNextUCase = "COMPILEERROREVENT" _
                   OrElse _
                   strNextUCase = "INTERPRETERROREVENT" Then
                    _OBJutilities.append(strOutstring, _
                                         vbNewLine, _
                                         Mid(strNext, 1, Len(strNext) - 5) & ": " & _
                                         CStr(colHandle.Item(3)))
                End If
            Next intIndex1
            Return (strOutstring)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Compile and run code...use preexisting results as available
    '
    '
    ' In immediate mode, this function returns the qbVariable result, or 
    ' Nothing on failure.
    '
    ' In compile and run mode, this function returns True on success and
    ' False on failure.
    '
    '
    Private Function executeCode_(ByVal strSourceCode As String, _
                                  ByVal enuSourceType As ENUsourceCodeType) As Object
        If OBJstate.usrState.objScanner.SourceCode <> strSourceCode Then
            OBJstate.usrState.objScanner.SourceCode = strSourceCode
        End If
        With OBJstate.usrState
            .enuSourceCodeType = enuSourceType
            If Not OBJstate.usrState.objScanner.Scanned Then
                If Not OBJstate.usrState.objScanner.scan Then
                    raiseEvent_("userErrorEvent", _
                                "Cannot scan this command (or nothing to do)", _
                                "Invalid characters have been found or the command " & _
                                "is empty")
                    clearScanner_(.objScanner)
                    Return (IIf(enuSourceType = ENUsourceCodeType.immediateCommand, Nothing, False))
                End If
            End If
            If Not OBJstate.usrState.booCompiled Then
                If Not compile_() Then
                    raiseEvent_("userErrorEvent", _
                                "Cannot parse and compile this code", "")
                    clearPolish_(.colPolish)
                    Return (IIf(enuSourceType = ENUsourceCodeType.immediateCommand, Nothing, False))
                End If
                If Not (.objImmediateResult Is Nothing) Then Return (.objImmediateResult)
                .booAssembled = False
            End If
            If Not .booAssembled Then
                If Not assemble_() Then
                    raiseEvent_("userErrorEvent", _
                                "Cannot assemble this code", "")
                    Return (IIf(enuSourceType = ENUsourceCodeType.immediateCommand, Nothing, False))
                End If
            End If
            If .enuSourceCodeType = ENUsourceCodeType.program OrElse (.objImmediateResult Is Nothing) Then
                interpret_()
            End If
            Return (IIf(enuSourceType = ENUsourceCodeType.immediateCommand, .objImmediateResult, True))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Extend the event log
    '
    '
    Private Function extendEventLog_(ByVal strEvent As String, _
                                     ByVal strOperands As String) As Boolean
        With OBJstate.usrState
            With .colEventLog
                Dim colEntry As Collection
                Try
                    colEntry = New Collection
                    With colEntry
                        .Add(strEvent)
                        .Add(Now)
                        .Add(strOperands)
                    End With
                    .Add(colEntry)
                Catch objException As Exception
                    errorHandler_("Cannot create the event log: " & _
                                  Err.Number & " " & Err.Description, _
                                  "extendEventLog_", _
                                  "Marking object unusable", _
                                  objException)
                    OBJstate.usrState.booUsable = False
                    Return (False)
                End Try
            End With
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Evaluate an immediate command
    '
    '
    Private Function evaluate_(ByVal strExpression As String, _
                               ByVal booEventLog As Boolean, _
                               ByRef strEventLog As String) As qbVariable.qbVariable
        Try
            Dim objQuickBasicEngine As quickBasicEngine.qbQuickBasicEngine = clone_()
            With objQuickBasicEngine
                .SourceCode = strExpression
                .ConstantFolding = False
                .EventLogging = booEventLog
                .run(ENUsourceCodeType.immediateCommand.ToString)
                If booEventLog Then strEventLog = .eventLogFormat
                Dim objResult As qbVariable.qbVariable = .Evaluation
                .dispose(OBJstate.usrState.booInspectCompilerObjects)
                Return (objResult)
            End With
        Catch objException As Exception
            errorHandler_("Error occured in the evaluation of " & _
                        _OBJutilities.enquote(_OBJutilities.ellipsis(strExpression, 32)) & ": " & _
                        Err.Number & " " & Err.Description, _
                        "evaluate", _
                        "Returning Nothing", _
                        objException)
            Return (Nothing)
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Event log formatting
    '
    '
    Private Shared Function eventLogFormat_(ByVal colEventLog As Collection, _
                                            ByVal intStartIndex As Integer, _
                                            ByVal intCount As Integer) _
           As String
        If intStartIndex < 1 OrElse intCount < 0 Then
            errorHandler_("Invalid start index or count", _
                          "eventLogFormat", _
                          "Returning a null string")
            Return ("")
        End If
        If (colEventLog Is Nothing) Then
            Return ("No event log exists")
        End If
        With colEventLog
            If .Count = 0 Then Return ("The event log is empty")
            ' --- Pass one: determine column max widths
            Dim colEntry As Collection
            Dim colSplit As Collection
            Try
                colSplit = New Collection
            Catch
                errorHandler_("Cannot create collection: " & _
                              Err.Number & " " & Err.Description, _
                              "eventLogFormat", _
                              "Returning a null string")
            End Try
            Dim colSplitInner1 As Collection
            Dim colSplitInner2 As Collection
            Dim intMaxWidth(3) As Integer
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim intIndex3 As Integer
            Dim strCol1Header As String = "Event Name"
            Dim strCol2Header As String = "Date and Time"
            Dim strCol3Header As String = "Operand"
            Dim strCol4Header As String = "O p e r a n d   V a l u e"
            intMaxWidth(0) = Len(strCol1Header)
            intMaxWidth(1) = Len(strCol2Header)
            intMaxWidth(2) = Len(strCol3Header)
            intMaxWidth(3) = Len(strCol4Header)
            Dim intSep(1) As Integer
            intSep(0) = 1 : intSep(1) = 1
            Dim strItem As String
            Dim strSplit() As String
            Dim intLength As Integer
            Dim intLimit As Integer = intStartIndex + intCount - 1
            For intIndex1 = intStartIndex To intLimit
                colEntry = CType(.Item(intIndex1), Collection)
                With colEntry
                    intMaxWidth(0) = Math.Max(intMaxWidth(0), Len(CStr(.Item(1))))
                    intMaxWidth(1) = Math.Max(intMaxWidth(1), Len(CStr(.Item(2))))
                    Try
                        strSplit = Split(CStr(.Item(3)), vbNewLine)
                        If UBound(strSplit) > 0 Then
                            intSep(0) = Math.Max(intSep(0), 2)
                        End If
                        colSplitInner1 = New Collection
                        For intIndex2 = 0 To UBound(strSplit)
                            colSplitInner2 = New Collection
                            With colSplitInner2
                                .Add(Trim(_OBJutilities.item(strSplit(intIndex2), 1, "=", False)))
                                intMaxWidth(2) = Math.Max(intMaxWidth(2), Len(CStr(.Item(1))))
                                strItem = Trim(_OBJutilities.item(strSplit(intIndex2), _
                                                                    2, _
                                                                    "=", _
                                                                    False))
                                .Add(Mid(strItem, 1, 40))
                                intMaxWidth(3) = Math.Min(40, Math.Max(intMaxWidth(3), Len(CStr(.Item(2)))))
                            End With
                            intLength = Len(strItem)
                            If intLength > 40 Then
                                intSep(1) = Math.Min(intSep(1) + 1, 1)
                                For intIndex3 = 41 To intLength Step 40
                                    colSplitInner2 = New Collection
                                    With colSplitInner2
                                        .Add("")
                                        .Add(Mid(strItem, intIndex3, 40))
                                    End With
                                Next intIndex3
                            End If
                            colSplitInner1.Add(colSplitInner2)
                        Next intIndex2
                        colSplit.Add(colSplitInner1)
                    Catch objException As Exception
                        errorHandler__("Cannot split/create collection: " & _
                                        Err.Number & " " & Err.Description, _
                                        ClassName, _
                                        "eventLogFormat", _
                                        "Returning a null string", _
                                        objException)
                    End Try
                End With
            Next intIndex1
            ' --- Pass two: create log
            Dim strOutstring As String
            Dim strIndent As String = _
                _OBJutilities.copies(" ", intMaxWidth(0) + intMaxWidth(1) + 2)
            For intIndex1 = intStartIndex To intLimit
                colEntry = CType(.Item(intIndex1), Collection)
                With colEntry
                    strItem = CStr(.Item(3))
                    colSplitInner1 = CType(colSplit.Item(intIndex1), Collection)
                    colSplitInner2 = CType(colSplitInner1.Item(1), Collection)
                    _OBJutilities.append(strOutstring, _
                                            _OBJutilities.copies(intSep(0), vbNewLine), _
                                            _OBJutilities.alignLeft(CStr(.Item(1)), intMaxWidth(0)) & " " & _
                                            _OBJutilities.alignLeft(CStr(.Item(2)), intMaxWidth(1)) & " " & _
                                            _OBJutilities.alignLeft(CStr(colSplitInner2.Item(1)), _
                                                                    intMaxWidth(2)) & " " & _
                                            _OBJutilities.alignLeft(CStr(colSplitInner2.Item(2)), _
                                                                    intMaxWidth(3)))
                    For intIndex2 = 2 To colSplitInner1.Count
                        colSplitInner2 = CType(colSplitInner1.Item(intIndex2), Collection)
                        _OBJutilities.append(strOutstring, _
                                            _OBJutilities.copies(intSep(1), vbNewLine), _
                                            strIndent & _
                                            _OBJutilities.alignLeft(CStr(colSplitInner2.Item(1)), _
                                                                    intMaxWidth(2)) & " " & _
                                            _OBJutilities.alignLeft(Replace(CStr(colSplitInner2.Item(2)), _
                                                                            vbNewLine, _
                                                                            "\n"), _
                                                                    intMaxWidth(3)))
                    Next intIndex2
                End With
            Next intIndex1
            Return (_OBJutilities.alignCenter(strCol1Header, intMaxWidth(0)) & " " & _
                    _OBJutilities.alignCenter(strCol2Header, intMaxWidth(1)) & " " & _
                    _OBJutilities.alignCenter(strCol3Header, intMaxWidth(2)) & " " & _
                    _OBJutilities.alignLeft(strCol4Header, intMaxWidth(3)) & " " & _
                    vbNewLine & _
                    _OBJutilities.copies("-", intMaxWidth(0)) & " " & _
                    _OBJutilities.copies("-", intMaxWidth(1)) & " " & _
                    _OBJutilities.copies("-", intMaxWidth(2)) & " " & _
                    _OBJutilities.copies("-", intMaxWidth(3)) & " " & _
                    vbNewLine & _
                    strOutstring)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Get the event logging status
    '
    '
    Private Function eventLoggingGet_() As Boolean
        Return Not (OBJstate.usrState.colEventLog Is Nothing)
    End Function

    ' -----------------------------------------------------------------
    ' Set and clear event logging
    '
    '
    Private Function eventLoggingSet_(ByVal booNewValue As Boolean) _
            As Boolean
        With OBJstate.usrState
            If booNewValue Then
                If (.colEventLog Is Nothing) Then
                    Try
                        .colEventLog = New Collection
                    Catch objException As Exception
                        errorHandler_("Can't create the event log", _
                                      "", _
                                      "Continuing: event log won't be provided", _
                                      objException)
                        Return (False)
                    End Try
                    Return (True)
                Else
                    Return (.objCollectionUtilities.collectionClear(.colEventLog))
                End If
            End If
        End With
    End Function

#If QUICKBASICENGINE_EXTENSION Then

    ' -----------------------------------------------------------------
    ' Get the event logging status
    '
    '
    Private Function generateNOPsGet_() As Boolean
        Return OBJstate.usrState.booGenerateNOPs
    End Function

    ' -----------------------------------------------------------------
    ' Get the event logging status
    '
    '
    Private Sub generateNOPsSet_(ByVal booNewValue As Boolean)
        OBJstate.usrState.booGenerateNOPs = booNewValue
    End Sub

#End If

    ' -----------------------------------------------------------------
    ' Inspect the object for internal errors
    '
    '
    Private Function inspect_(ByRef strReport As String) As Boolean
        With OBJstate.usrState
            Dim booInspection As Boolean = True
            If _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_USABLE, _
                                            .booUsable, _
                                            booInspection) Then
                Dim strSubreport As String
                _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_SCANNEROK, _
                                            (.objScanner Is Nothing) _
                                            OrElse _
                                            .objScanner.inspect(strSubreport), _
                                            booInspection, _
                                            strSubreport, _
                                            booString2Box:=False)
                _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_POLISHOK, _
                                            (.colPolish Is Nothing) _
                                            OrElse _
                                            inspect_collection_(.colPolish, strSubreport), _
                                            booInspection, _
                                            strSubreport, _
                                            booString2Box:=False)
                _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_VARIABLESOK, _
                                            (.colVariables Is Nothing) _
                                            OrElse _
                                            inspect_collection_(.colVariables, strSubreport), _
                                            booInspection, _
                                            strSubreport, _
                                            booString2Box:=False)
                _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_SUBFUNCTIONINDEX, _
                                            (.colSubFunctionIndex Is Nothing) _
                                            OrElse _
                                            inspect_index_(.colSubFunctionIndex, _
                                                            strSubreport, _
                                                            UBound(.usrSubFunction)), _
                                            booInspection, _
                                            strSubreport, _
                                            booString2Box:=False)
            End If
            If Not booInspection Then
                OBJstate.usrState.booUsable = False : Return (False)
            End If
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the collection randomly or exhaustively on behalf
    ' of inspect
    '
    '
    Private Function inspect_collection_(ByVal colCollection As Collection, _
                                         ByRef strReport As String) As Boolean
        Dim booPolish As Boolean
        Dim booInspect As Boolean
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim objPolishHandle As qbPolish.qbPolish
        Dim objVariableHandle As qbVariable.qbVariable
        strReport = ""
        With colCollection
            If .Count = 0 Then
                strReport = "Collection is empty" : Return (True)
            End If
            If .Count > INSPECTION_MAXCOLLECTIONITEMS Then
                intCount = CInt(Rnd() * INSPECTION_MAXCOLLECTIONITEMS)
            Else
                intCount = .Count
            End If
            For intIndex1 = 1 To intCount
                intIndex2 = intIndex1
                If .Count > INSPECTION_MAXCOLLECTIONITEMS Then
                    intIndex2 = Math.Min(.Count, CInt(Rnd() * INSPECTION_MAXCOLLECTIONITEMS) + 1)
                End If
                objPolishHandle = Nothing : objVariableHandle = Nothing
                Try
                    objPolishHandle = CType(.Item(intIndex2), qbPolish.qbPolish)
                Catch
                    Try
                        objVariableHandle = CType(.Item(intIndex2), qbVariable.qbVariable)
                    Catch objException As Exception
                        strReport = "Unsupported object type"
                        Exit For
                    End Try
                End Try
                If intIndex1 = 1 Then
                    booPolish = Not (objPolishHandle Is Nothing)
                ElseIf booPolish And (objPolishHandle Is Nothing) _
                       OrElse _
                       Not booPolish And (objVariableHandle Is Nothing) Then
                    strReport = "Inconsistent mix of object types occurs"
                    Exit For
                End If
                If Not (objPolishHandle Is Nothing) Then
                    booInspect = objPolishHandle.inspect(strReport)
                Else
                    booInspect = objVariableHandle.inspect(strReport)
                    Try
                        Dim objHandle As qbVariable.qbVariable = _
                            CType(.Item(objVariableHandle.VariableName), _
                                  qbVariable.qbVariable)
                        Dim intTag As Integer = CInt(objHandle.Tag)
                        If intTag <> intIndex1 Then
                            strReport = "At variable collection item " & intIndex1 & ", " & _
                                        "an invalid Tag occurs"
                        End If
                    Catch
                        strReport = "At variable collection item " & intIndex1 & " " & _
                                    "the variable fails to properly index itself: " & _
                                    Err.Number & " " & Err.Description
                        Exit For
                    End Try
                End If
                If Not booInspect Then
                    strReport = "In the " & .Item(intIndex1).GetType.ToString & " " & _
                                "collection, object " & _
                                intIndex2 & " of " & .Count & " " & _
                                "fails its own inspection:" & _
                                vbNewLine & vbNewLine & _
                                strReport
                    Exit For
                End If
                If loopEventInterface_("Inspecting " & .Item(intIndex1).GetType.ToString & " objects", _
                                       .Item(intIndex1).GetType.ToString & " object", _
                                       intIndex1, intCount, _
                                       0, _
                                       "") Then Exit For
            Next intIndex1
            If intIndex1 <= intCount Then Return (False)
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the subroutine/function index on behalf of inspect
    '
    '
    Private Function inspect_index_(ByVal colIndex As Collection, _
                                    ByRef strReport As String, _
                                    ByVal intUBound As Integer) As Boolean
        Dim colHandle As Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        strReport = ""
        With colIndex
            For intIndex1 = 1 To .Count
                Try
                    colHandle = CType(.Item(intIndex1), Collection)
                Catch
                    strReport = "Entry is not a collection at " & intIndex1
                    Exit For
                End Try
                With colHandle
                    Try
                        Dim strName As String = CStr(.Item(1))
                        intIndex2 = CInt(.Item(2))
                    Catch
                        strReport = "Entry doesn't contain expected fields at " & intIndex1
                        Exit For
                    End Try
                    If intIndex2 < 1 OrElse intIndex2 > intUBound Then
                        strReport = "Entry contains invalid constant expression index at " & intIndex1
                        Exit For
                    End If
                End With
                If loopEventInterface_("Inspecting constant expression index", _
                                       "item", _
                                       intIndex1, _
                                       .Count, _
                                       0, _
                                       "") Then Exit For
            Next intIndex1
        End With
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the subroutine/function index on behalf of inspect
    '
    '
    Private Function inspect_subFunctionIndex_(ByRef strReport As String) As Boolean
        Dim colHandle As Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        strReport = ""
        With OBJstate.usrState.colSubFunctionIndex
            For intIndex1 = 1 To .Count
                Try
                    colHandle = CType(.Item(intIndex1), Collection)
                Catch
                    strReport = "Entry is not a collection at " & intIndex1
                    Exit For
                End Try
                With colHandle
                    Try
                        Dim strName As String = CStr(.Item(1))
                        intIndex2 = CInt(.Item(2))
                    Catch
                        strReport = "Entry doesn't contain expected fields at " & intIndex1
                        Exit For
                    End Try
                    If intIndex2 < 1 _
                       OrElse _
                       intIndex2 > UBound(OBJstate.usrState.usrSubFunction) Then
                        strReport = "Entry contains invalid subroutine/function index at " & intIndex1
                        Exit For
                    End If
                End With
                If loopEventInterface_("Inspecting subroutine/function index", _
                                       "item", _
                                       intIndex1, _
                                       .Count, _
                                       0, _
                                       "") Then Exit For
            Next intIndex1
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Internal inspection
    '
    '
    Private Function inspection_() As Boolean
        Dim strReport As String
        If Not inspect_(strReport) Then
            errorHandler_("Internal inspection failed" & vbNewLine & vbNewLine & _
                          "inspection_", _
                          "Object is not usable", _
                          Nothing)
            Return (False)
        End If
        Return (True)
    End Function

#Region "Interpreter"

    ' ----------------------------------------------------------------------
    ' Interpret
    '
    '
    Private Function interpret_() As Object
        With OBJstate.usrState
            Try
                Dim objResult As Object = _
                    interpreter_(.colPolish, _
                                    .enuSourceCodeType _
                                    = _
                                    ENUsourceCodeType.immediateCommand)
                If (TypeOf objResult Is qbVariable.qbVariable) Then
                    .objImmediateResult = CType(objResult, qbVariable.qbVariable)
                End If
                Return (objResult)
            Catch objException As Exception
                errorHandler_("Error occured in interpreter: " & _
                            Err.Number & " " & Err.Description, _
                            "interpret_", _
                            "Marking object unusable and returning Nothing", _
                            objException)
                .booUsable = False : Return (Nothing)
            End Try
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Interpret the compiled code: return expression value as a qbVariable 
    ' (for an immediate command) or True/False success 
    '
    '
    ' Note that the interpreter has two different modes depending on the
    ' setting of the compile-time symbol QUICKBASICENGINE_POPCHECK:
    '
    '
    '      *  If QUICKBASICENGINE_POPCHECK is True then the interpreter
    '         uses the objStack and objFrame.
    '
    '         objStack is a .Net Stack collection restricted by code to
    '         qbVariables; objFrame is an array also restricted to
    '         qbVariables. 
    '
    '         Prior to the execution of each instruction, the expected
    '         stack frame associated with the instruction is popped and
    '         checked for validity.
    '
    '      *  If QUICKBASICENGINE_POPCHECK is False then the interpreter
    '         uses the objStack which here is a qbVariable array, along with
    '         intStackTop.
    '
    '         Stack frames are not checked for validity and the frame copy
    '         isn't performed which results in a speed advantage at the cost
    '         of less internal checking. 
    '
    '
    Private Function interpreter_(ByRef colPolish As Collection, _
                                  ByVal booImmediate As Boolean) As Object
        Dim booAnyPrinting As Boolean
        Dim booNumeric As Boolean
        Dim booOK As Boolean
        Dim booTraceThisLine As Boolean
        Dim dblValue As Double
        Dim intCount As Integer
        Dim intIndexIP As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intLastMemReference As Integer
        Dim intLength As Integer
        Dim intTest As Integer
        Dim intUBound As Integer
        Dim objHandle As qbVariable.qbVariable
        Dim objInputQueue As Queue              ' Of type qbvariable
        Dim objQuickBasicEngine As New quickBasicEngine.qbQuickBasicEngine
        Dim objTraceStack As Stack
        Dim objValue As Object
        Dim usrTracing As TYPtracing
        Dim objVariable As qbVariable.qbVariable
#If QUICKBASICENGINE_POPCHECK Then
        Dim objFrame() As qbVariable.qbVariable
        Dim objStack As Stack  ' Of qbVariables    
#Else
        Dim dblStep As Double
        Dim dblLast As Double
        Dim objOperand1 As qbVariable.qbVariable
        Dim objOperand2 As qbVariable.qbVariable
        Dim objOperand3 As qbVariable.qbVariable
        Dim objStack() As qbVariable.qbVariable
        Dim objStackCopy As Stack
        Dim intStackTop As Integer
#End If
        Try
#If QUICKBASICENGINE_POPCHECK Then
            objStack = New Stack
#End If
            objTraceStack = New Stack
            objInputQueue = New Queue
        Catch objException As Exception
            errorHandler_("Cannot allocate the interpreter data structures: " & _
                          Err.Number & " " & Err.Description, _
                          "interpreter_", _
                          "Making this object unusable: returning Nothing", _
                          objException)
            OBJstate.usrState.booUsable = False
            Return (Nothing)
        End Try
#If QUICKBASICENGINE_POPCHECK Then
        If Not interpreter__expandFrame_(objFrame, 0) Then Return Nothing
#Else
        Try
            ReDim objStack(INTERPRETER_STACK_BLOCK)
        Catch objException As Exception
            errorHandler_("Cannot allocate the ""fast"" stack: " & _
                          Err.Number & " " & Err.Description, _
                          "interpreter_", _
                          "Making this object unusable: returning Nothing", _
                          objException)
            OBJstate.usrState.booUsable = False
            Return (Nothing)
        End Try
#End If
        clearStorage_()
        intIndexIP = 1 : intLastMemReference = 1
        objTraceStack.Push(0)
        OBJstate.usrState.intReadDataIndex = 1
        intCount = 0
        Do While intIndexIP <= colPolish.Count
            If OBJthreadStatus.getThreadStatus = OBJthreadStatus.ENUthreadStatus.Stopping Then Exit Do
            booTraceThisLine = False
            If objTraceStack.Count > 1 Then
                usrTracing = CType(traceCode2Struct_(CInt(objTraceStack.Peek)), TYPtracing)
                booTraceThisLine = usrTracing.intRate <> 0 AndAlso intCount Mod usrTracing.intRate = 0 _
                                   AndAlso _
                                   (Not usrTracing.booTraceLines _
                                    OrElse _
                                    intIndexIP = 1 _
                                    OrElse _
                                    interpreter__polish2Source_(intIndexIP) _
                                    <> _
                                    interpreter__polish2Source_(intIndexIP))
            End If
#If QUICKBASICENGINE_POPCHECK Then
            raiseEvent_("interpretTraceEvent", _
                        intIndexIP, _
                        objStack, _
                        OBJstate.usrState.colVariables)
#Else
            RaiseEvent interpretTraceFastEvent(Me, _
                                               intIndexIP, _
                                               objStack, _
                                               intStackTop, _
                                               OBJstate.usrState.colVariables)
#End If
            With CType(colPolish.Item(intIndexIP), qbPolish.qbPolish)
                Try
#If QUICKBASICENGINE_POPCHECK Then
                    If Not interpreter__pop_(objStack, _
                                             _OBJqbOp.opCodeToStackTemplate(.Opcode), _
                                             objFrame) Then
                        Exit Do
                    End If
#End If
                    Select Case .Opcode
                        Case ENUop.opAdd
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CDbl(objFrame(1).value) _
                                               + _
                                               CDbl(objFrame(2).value))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(objOperand1.value) _
                                               + _
                                               CDbl(objOperand2.value))
#End If
                        Case ENUop.opAnd
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CInt(IIf(CBool(objFrame(1).value) _
                                                        AndAlso _
                                                        CBool(objFrame(2).value), _
                                                        -1, 0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CBool(objOperand1.value) _
                                               AndAlso _
                                               CBool(objOperand2.value))
#End If
                        Case ENUop.opAsc
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Asc(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Asc(CStr(interpreter__pop_(objStack, intStackTop).value)))
#End If

                        Case ENUop.opCeil
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Ceiling(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Ceiling(CDbl(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opChr
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, ChrW(CInt(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               ChrW(CInt(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opCircle
                            interpreter__err_("The Circle operation hasn't been implemented yet")
                        Case ENUop.opCls
                            raiseEvent_("interpretClsEvent")
                        Case ENUop.opCoGo
#If QUICKBASICENGINE_POPCHECK Then
                            Dim strLabel As String = CStr(objFrame(1).value)
#Else
                            Dim strLabel As String = CStr(interpreter__pop_(objStack, intStackTop).value)
#End If
                            Try
                                intIndex1 = _
                                CInt(CType(OBJstate.usrState.colLabel.Item("_" & strLabel), _
                                           Collection).Item(2))
                            Catch
#If QUICKBASICENGINE_POPCHECK Then
                                interpreter__err_("Error in converting the label " & _
                                                    strLabel & " " & _
                                                    "to an address: " & _
                                                    Err.Number & " " & Err.Description, _
                                                    intIndexIP, _
                                                    objStack, _
                                                    "")
#Else
                                interpreter__err_("Error in converting the label " & _
                                                    strLabel & " " & _
                                                    "to an address: " & _
                                                    Err.Number & " " & Err.Description)
#End If
                                Exit Do
                            End Try
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, intIndex1)
#Else
                            interpreter__push_(objStack, intStackTop, intIndex1)
#End If
                        Case ENUop.opConcat
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CStr(objFrame(1).value) _
                                               & _
                                               CStr(objFrame(2).value))
#If not QUICKBASICENGINE_EXTENSION Then
                            objHandle = CType(objStack.peek, _ 
                                              qbVariable.qbVariable)   
#End If
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CStr(objOperand1.value) _
                                               & _
                                               CStr(objOperand2.value))
#If Not QUICKBASICENGINE_EXTENSION Then
                            objHandle = objStack(intStackTop)
#End If
#End If
#If Not QUICKBASICENGINE_EXTENSION Then
                            interpreter__stringTruncate_(objHandle)
#End If
                        Case ENUop.opCos
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Cos(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Cos(CDbl(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opDivide
#If QUICKBASICENGINE_POPCHECK Then
                            Dim dblDivisor As Double = CDbl(objFrame(2).value)
                            If dblDivisor = 0 Then
                                interpreter__err_("Division by zero")
                            End If
                            interpreter__push_(objStack, CDbl(objFrame(1).value) / dblDivisor)
#Else
                            Dim dblDivisor As Double = CDbl(interpreter__pop_(objStack, intStackTop).value)
                            If dblDivisor = 0 Then
                                interpreter__err_("Division by zero")
                            End If
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(interpreter__pop_(objStack, intStackTop).value) _
                                               / _
                                               dblDivisor)
#End If
                        Case ENUop.opDuplicate
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, objStack.Peek)
#Else
                            interpreter__push_(objStack, intStackTop, objStack(intStackTop))
#End If
                        Case ENUop.opEnd
                            Exit Do
                        Case ENUop.opEval
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, eval(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               eval(CStr(interpreter__pop_(objStack, _
                                                                            intStackTop).value)))
#End If
                        Case ENUop.opEvaluate
                            Dim strDummy As String
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               evaluate_(CStr(objFrame(1).value), False, strDummy))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               evaluate_(CStr(interpreter__pop_(objStack, intStackTop).value), _
                                                         False, _
                                                         strDummy))
#End If
                        Case ENUop.opFloor
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Floor(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Floor(CDbl(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opForIncrement
                            ' Stack frame: step, index location
#If QUICKBASICENGINE_POPCHECK Then
                            intIndex1 = CInt(objFrame(2).value)
                            objVariable = CType(OBJstate.usrState.colVariables(intIndex1), _
                                                qbVariable.qbVariable)
                            objVariable.valueSet(CDbl(objVariable.value) + CDbl(objFrame(1).value))
                            interpreter__push_(objStack, objFrame(1))
                            interpreter__push_(objStack, objFrame(2))
#Else
                            objVariable = _
                            CType(OBJstate.usrState.colVariables.Item(CInt(objStack(intStackTop).value)), _
                                  qbVariable.qbVariable)
                            objVariable.valueSet(CDbl(objVariable.value) + CDbl(objStack(intStackTop - 1).value))
#End If
                        Case ENUop.opForTest
                            ' Stack frame: last, step, index location
#If QUICKBASICENGINE_POPCHECK Then
                            dblValue = CDbl(CType(OBJstate.usrState.colVariables(CInt(objFrame(3).value)), _
                                                  qbVariable.qbVariable).value)
                            If CDbl(objFrame(2).value) >= 0 _
                               AndAlso _
                               dblValue > CDbl(objFrame(1).value) _
                               OrElse _
                               CDbl(objFrame(2).value) < 0 _
                               AndAlso _
                               dblValue < CDbl(objFrame(1).value) Then
                                intIndexIP = CInt(.Operand) - 1
                            End If
                            interpreter__push_(objStack, objFrame(1))
                            interpreter__push_(objStack, objFrame(2))
                            interpreter__push_(objStack, objFrame(3))
#Else
                            dblValue = CDbl(CType(OBJstate.usrState.colVariables(CInt(objStack(intStackTop).value)), _
                                                  qbVariable.qbVariable).value)
                            dblStep = CDbl(objStack(intStackTop - 1).value)
                            dblLast = CDbl(objStack(intStackTop - 2).value)
                            If dblStep >= 0 AndAlso dblValue > dblLast _
                               OrElse _
                               dblStep < 0 AndAlso dblValue < dblLast Then
                                intIndexIP = CInt(.Operand) - 1
                            End If
#End If
                        Case ENUop.opIif
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               IIf(CDbl(objFrame(1).value) <> 0, _
                                                   objFrame(2).value, _
                                                   objFrame(3).value))
#Else
                            objOperand3 = interpreter__pop_(objStack, intStackTop)
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               IIf(CDbl(objOperand1.value) <> 0, _
                                                   objOperand2, _
                                                   objOperand1))
#End If
                        Case ENUop.opInput
                            Dim objInput As qbVariable.qbVariable = _
                                interpreter__input_(objInputQueue)
                            If (objInput Is Nothing) Then Exit Do
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, objInput)
#Else
                            interpreter__push_(objStack, intStackTop, objInput)
#End If
                        Case ENUop.opInt
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, CInt(objFrame(1).value))
#Else
                            If Not objStack(intStackTop).isAnUnsignedInteger Then
                                interpreter__push_(objStack, _
                                                   intStackTop, _
                                                   CInt(interpreter__pop_(objStack, intStackTop).value))
                            End If
#End If
                        Case ENUop.opIsNumeric
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, objFrame(1).isaNumber)
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               interpreter__pop_(objStack, intStackTop).isaNumber)
#End If
                        Case ENUop.opJump
                            intIndexIP = CInt(.Operand) - 1
                        Case ENUop.opJumpIndirect
#If QUICKBASICENGINE_POPCHECK Then
                            intIndexIP = CInt(objFrame(1).value) - 1
#Else
                            intIndexIP = CInt(interpreter__pop_(objStack, intStackTop).value) - 1
#End If
                        Case ENUop.opJumpNZ
#If QUICKBASICENGINE_POPCHECK Then
                            If CDbl(objFrame(1).value) <> 0 Then intIndexIP = CInt(.Operand) - 1
#Else
                            If CDbl(interpreter__pop_(objStack, intStackTop).value) <> 0 Then
                                intIndexIP = CInt(.Operand) - 1
                            End If
#End If
                        Case ENUop.opJumpZ
#If QUICKBASICENGINE_POPCHECK Then
                            If CDbl(objFrame(1).value) = 0 Then intIndexIP = CInt(.Operand) - 1
#Else
                            If CDbl(interpreter__pop_(objStack, intStackTop).value) = 0 Then
                                intIndexIP = CInt(.Operand) - 1
                            End If
#End If
                        Case ENUop.opLabel
                        Case ENUop.opLCase
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, LCase(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               LCase(CStr(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opLen
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, LCase(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Len(CStr(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opLike
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, UCase(CStr(objFrame(1).value)) Like UCase(CStr(objFrame(2).value)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CStr(objOperand1.value) Like CStr(objOperand2.value))
#End If
                        Case ENUop.opLog
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Log(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Log(CDbl(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opMax
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               Math.Max(CDbl(objFrame(1).value), CDbl(objFrame(2).value)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Max(CDbl(objOperand1.value), CDbl(objOperand2.value)))
#End If
                        Case ENUop.opMid
#If QUICKBASICENGINE_POPCHECK Then
                            If UBound(objFrame) = 3 Then
                                interpreter__push_(objStack, Mid(CStr(objFrame(1).value), _
                                                                 CInt(objFrame(2).value), _
                                                                 CInt(objFrame(3).value)))
                            Else
                                interpreter__push_(objStack, Mid(CStr(objFrame(1).value), _
                                                                 CInt(objFrame(2).value)))
                            End If
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            If intStackTop >= 1 Then
                                interpreter__push_(objStack, _
                                                   intStackTop, _
                                                   Mid(CStr(objOperand1.value), _
                                                       CInt(objOperand2.value), _
                                                       CInt(interpreter__pop_(objStack, intStackTop).value)))
                            Else
                                interpreter__push_(objStack, _
                                                   intStackTop, _
                                                   Mid(CStr(objOperand1.value), _
                                                       CInt(objOperand2.value)))
                            End If
#End If
                        Case ENUop.opMin
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               Math.Min(CDbl(objFrame(1).value), CDbl(objFrame(2).value)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Min(CDbl(objOperand1.value), CDbl(objOperand2.value)))
#End If
                        Case ENUop.opMod
#If QUICKBASICENGINE_POPCHECK Then
                            Dim dblDivisorMod As Double = CDbl(objFrame(2).value)
                            If dblDivisorMod = 0 Then
                                interpreter__err_("Division by zero in modulus operation", _
                                                    intIndexIP, _
                                                    objStack, _
                                                    "")
                                Exit Do
                            End If
                            interpreter__push_(objStack, _
                                               CDbl(objFrame(1).value) Mod dblDivisorMod)
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(objOperand1.value) Mod CDbl(objOperand2.value))
#End If
                        Case ENUop.opMultiply
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CDbl(objFrame(1).value) _
                                               * _
                                               CDbl(objFrame(2).value))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(objOperand1.value) * CDbl(objOperand2.value))
#End If
                        Case ENUop.opNegate
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, -CDbl(objFrame(1).value))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               -CDbl(interpreter__pop_(objStack, intStackTop).value))
#End If
                        Case ENUop.opNop
                        Case ENUop.opNot
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CInt(IIf(CBool(objFrame(1).value), _
                                                        0, -1)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(CBool(interpreter__pop_(objStack, _
                                                                                intStackTop).value), _
                                                        0, -1)))
#End If
                        Case ENUop.opOr
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CInt(IIf(CBool(objFrame(1).value) _
                                                        OrElse _
                                                        CBool(objFrame(2).value), _
                                                        -1, 0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CBool(objOperand1.value) OrElse CBool(objOperand2.value))
#End If
                        Case ENUop.opPop
                            objHandle = _
                                CType(OBJstate.usrState.colVariables.Item(CInt(.Operand)), _
                                      qbVariable.qbVariable)
#If QUICKBASICENGINE_POPCHECK Then
                            objHandle.valueSet(objFrame(1).value)
#Else
                            objHandle.valueSet(interpreter__pop_(objStack, intStackTop).value)
#End If
                        Case ENUop.opPopIndirect
#If QUICKBASICENGINE_POPCHECK Then
                            Dim intIndirect As Integer = CInt(objFrame(1).value)
                            CType(OBJstate.usrState.colVariables.Item(intIndirect), _
                                  qbVariable.qbVariable).valueSet _
                            (objFrame(2).value)
                            interpreter__push_(objStack, intIndirect)
#Else
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            CType(OBJstate.usrState.colVariables.Item(objStack(intStackTop).value), _
                                  qbVariable.qbVariable).valueSet _
                            (objOperand1.value)
#End If
                        Case ENUop.opPopOff
#If Not QUICKBASICENGINE_POPCHECK Then
                            interpreter__pop_(objStack, intStackTop)
#End If
                        Case ENUop.opPopToArrayElement
#If QUICKBASICENGINE_POPCHECK Then
                            CType(OBJstate.usrState.colVariables(CInt(objFrame(UBound(objFrame) - 1).value)), _
                                  qbVariable.qbVariable).valueSet _
                                  (objFrame(UBound(objFrame)), _
                                   interpreter__getSubscripts_(objFrame, 1, UBound(objFrame) - 1))
#Else
                            intIndex1 = intStackTop - 2
                            objStack(intStackTop - 1).valueSet _
                                  (objStack(intStackTop), _
                                   interpreter__getSubscripts_(objStack, _
                                                               intIndex1 _
                                                               - _
                                                               CInt(objStack(intIndex1).value), _
                                                               intIndex1 - 1))
                            interpreter__pop_(objStack, intStackTop, CInt(objStack(intIndex1).value) + 3)
#End If
                        Case ENUop.opPower
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               CDbl(objFrame(1).value) ^ CDbl(objFrame(2).value))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(objOperand1.value) ^ CDbl(objOperand2.value))
#End If
                        Case ENUop.opPrint
#If QUICKBASICENGINE_POPCHECK Then
                            raiseEvent_("interpretPrintEvent", CStr(objFrame(1).value))
#Else
                            raiseEvent_("interpretPrintEvent", _
                                        CStr(interpreter__pop_(objStack, intStackTop).value))
#End If
                        Case ENUop.opPush
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, .Operand)
#Else
                            interpreter__push_(objStack, intStackTop, .Operand)
#End If
                        Case ENUop.opPushArrayElement
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               _OBJqbVariable.mkVariable _
                                               (CStr(CType(objFrame(UBound(objFrame)), _
                                                           qbVariable.qbVariable).value _
                                                           (interpreter__getSubscripts_ _
                                                            (objFrame, _
                                                             UBound(objFrame) _
                                                             - _
                                                             CInt(objFrame(UBound(objFrame) - 1).value), _
                                                             UBound(objFrame) - 1)))))
#Else
                            intIndex1 = CInt(objStack(intStackTop - 2).value)
                            objValue = _OBJqbVariable.mkVariable _
                                        (CStr(objStack(intStackTop).value _
                                                    (interpreter__getSubscripts_ _
                                                    (objStack, _
                                                        intStackTop _
                                                        - _
                                                        CInt(objStack(intStackTop - 1).value), _
                                                        intStackTop - 1))))
                            interpreter__pop_(objStack, intStackTop, intIndex1 + 2)
                            interpreter__push_(objStack, intStackTop, objValue)
#End If
                        Case ENUop.opPushEQ
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        = _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        = _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        = _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        = _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushGE
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        >= _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        >= _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        >= _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        >= _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushGT
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        >  _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        >  _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        > _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        > _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushIndirect
#If QUICKBASICENGINE_POPCHECK Then
                            objValue = OBJstate.usrState.colVariables.Item(CInt(objFrame(1).value))
                            interpreter__push_(objStack, _
                                               objValue, _
                                               Not CType(objValue, qbVariable.qbVariable).Dope.isArray _
                                               AndAlso _
                                               Not CType(objValue, qbVariable.qbVariable).Dope.isUDT)
#Else
                            objValue = OBJstate.usrState.colVariables.Item _
                                       (CInt(interpreter__pop_(objStack, intStackTop).value))
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               objValue, _
                                               Not CType(objValue, qbVariable.qbVariable).Dope.isArray _
                                               AndAlso _
                                               Not CType(objValue, qbVariable.qbVariable).Dope.isUDT)
#End If
                        Case ENUop.opPushLE
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        <=  _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        <=  _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        <= _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        <= _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushLiteral
                            objValue = .Operand
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, objValue)
#Else
                            interpreter__push_(objStack, intStackTop, objValue)
#End If
                        Case ENUop.opPushLT
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        <  _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        <  _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        < _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        < _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushNE
#If QUICKBASICENGINE_POPCHECK Then
                            booNumeric = interpreter__resolveTypes_(objFrame(1), objFrame(2))
                            interpreter__push_(objStack, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objFrame(1).value) _
                                                        <>  _
                                                        CDbl(objFrame(2).value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objFrame(1).value) _
                                                        <>  _
                                                        CStr(objFrame(2).value), _
                                                        -1, _
                                                        0)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            booNumeric = interpreter__resolveTypes_(objOperand1, objOperand2)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               CInt(IIf(booNumeric _
                                                        AndAlso _
                                                        CDbl(objOperand1.value) _
                                                        <> _
                                                        CDbl(objOperand2.value) _
                                                        OrElse _
                                                        Not booNumeric _
                                                        AndAlso _
                                                        CStr(objOperand1.value) _
                                                        <> _
                                                        CStr(objOperand2.value), _
                                                        -1, _
                                                        0)))
#End If
                        Case ENUop.opPushReturn
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, .Operand)
#Else
                            interpreter__push_(objStack, intStackTop, .Operand)
#End If
                        Case ENUop.opRand
                            Randomize()
                        Case ENUop.opRead
                            With OBJstate.usrState
                                If .intReadDataIndex > .colReadData.Count Then
#If QUICKBASICENGINE_POPCHECK Then
                                    interpreter__err_("Read statement has no data", _
                                                      intIndexIP, objStack, _
                                                      "")
#Else
                                    interpreter__err_("Read statement has no data")
#End If
                                    Exit Do
                                End If
#If QUICKBASICENGINE_POPCHECK Then
                                interpreter__push_(objStack, .colReadData.Item(.intReadDataIndex))
#Else
                                interpreter__push_(objStack, _
                                                   intStackTop, _
                                                   .colReadData.Item(.intReadDataIndex))
#End If
                                .intReadDataIndex += 1
                            End With
                        Case ENUop.opRem
                        Case ENUop.opReplace
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               Replace(CStr(objFrame(1).value), _
                                                       CStr(objFrame(2).value), _
                                                       CStr(objFrame(3).value)))
#Else
                            objOperand3 = interpreter__pop_(objStack, intStackTop)
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Replace(CStr(objOperand1.value), _
                                                       CStr(objOperand2.value), _
                                                       CStr(objOperand3.value)))
#End If
                        Case ENUop.opRnd
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Rnd)
#Else
                            interpreter__push_(objStack, intStackTop, Rnd)
#End If
                        Case ENUop.opRndSeed
                            Rnd(CSng(.Operand))
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Rnd)
#Else
                            interpreter__push_(objStack, intStackTop, Rnd)
#End If
                        Case ENUop.opRotate
                            Dim intRotate As Integer = CInt(.Operand)
#If QUICKBASICENGINE_POPCHECK Then
                            If objStack.Count < intRotate + 1 Then
                                errorHandler_("Stack underflow in rotate: rotate value is " & _
                                              intRotate & ": " & _
                                              "stack size is " & objStack.Count, _
                                              "interpreter_", _
                                              "", _
                                              Nothing)
                                Exit Do
                            End If
                            If intRotate <> 0 Then
                                interpreter__pop_(objStack, _
                                                  _OBJutilities.copies("x,", intRotate) & "x", _
                                                  objFrame)
                                Dim objRotate1 As qbVariable.qbVariable = objFrame(1)
                                objFrame(1) = objFrame(UBound(objFrame))
                                objFrame(UBound(objFrame)) = objRotate1
                                For intIndex1 = 1 To UBound(objFrame)
                                    interpreter__push_(objStack, objFrame(intIndex1))
                                    disposeVariable_(objFrame(intIndex1))
                                    objFrame(intIndex1) = Nothing
                                Next intIndex1
                                ReDim objFrame(0)
                            End If
#Else
                            If intStackTop < intRotate + 1 Then
                                errorHandler_("Stack underflow in rotate: rotate value is " & _
                                              intRotate & ": " & _
                                              "stack size is " & intStackTop, _
                                              "interpreter_", _
                                              "", _
                                              Nothing)
                                Exit Do
                            End If
                            If intRotate <> 0 Then
                                intIndex1 = intStackTop - intRotate
                                Dim objRotate1 As qbVariable.qbVariable = objStack(intIndex1)
                                objStack(intIndex1) = objStack(intStackTop)
                                objStack(intStackTop) = objRotate1
                            End If
#End If
                        Case ENUop.opRound
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, _
                                               Math.Round(CDbl(objFrame(1).value), _
                                                          CInt(objFrame(2).value)))
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Round(CDbl(objOperand1.value), _
                                                          CInt(objOperand2.value)))
#End If
                        Case ENUop.opSgn
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Sign(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Sign(CDbl(interpreter__pop_(objStack, _
                                                                                intStackTop).value)))
#End If
                        Case ENUop.opSin
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Sin(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Sin(CDbl(interpreter__pop_(objStack, _
                                                                               intStackTop).value)))
#End If
                        Case ENUop.opSqr
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Math.Sqrt(CDbl(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Math.Sqrt(CDbl(interpreter__pop_(objStack, _
                                                                                intStackTop).value)))
#End If
                        Case ENUop.opString
#If QUICKBASICENGINE_POPCHECK Then
                            If Len(CStr(objFrame(1).value)) <> 1 Then
                                interpreter__err_("String function requires a single character string arg", _
                                                    intIndexIP, _
                                                    objStack, _
                                                    "")
                                Exit Do
                            End If
                            If CInt(objFrame(2).value) < 0 Then
                                interpreter__err_("String function requires count of zero or greater", _
                                                    intIndexIP, _
                                                    objStack, _
                                                    "")
                                Exit Do
                            End If
                            interpreter__push_(objStack, _
                                               _OBJutilities.copies(CInt(objFrame(2).value), _
                                                                    CStr(objFrame(1).value)))
#if Not QUICKBASICENGINE_EXTENSION then
                            objHandle = CType(objStack.peek, qbVariable.qbVariable)
#End If
#Else
                            objOperand2 = interpreter__pop_(objStack, intStackTop)
                            objOperand1 = interpreter__pop_(objStack, intStackTop)
                            If Len(CStr(objOperand1.value)) <> 1 Then
                                interpreter__err_("String function requires a single character string arg")
                                Exit Do
                            End If
                            If CInt(objOperand2.value) < 0 Then
                                interpreter__err_("String function requires count of zero or greater")
                                Exit Do
                            End If
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               _OBJutilities.copies(CInt(objOperand2.value), _
                                                                    CStr(objOperand1.value)))
#If Not QUICKBASICENGINE_EXTENSION Then
                            objHandle = CType(objStack(intstacktop), qbVariable.qbVariable)
#End If
#End If
#If Not QUICKBASICENGINE_EXTENSION Then
                            interpreter__stringTruncate_(objHandle)
#End If
                        Case ENUop.opSubtract
#If QUICKBASICENGINE_POPCHECK Then
                           interpreter__push_(objStack, _
                                               CDbl(objFrame(1).value) _
                                               - _
                                               CDbl(objFrame(2).value))
#Else
                           objOperand2 = interpreter__pop_(objStack, intStackTop)
                           objOperand1 = interpreter__pop_(objStack, intStackTop)
                           interpreter__push_(objStack, _
                                               intStackTop, _
                                               CDbl(objOperand1.value) _
                                               - _
                                               CDbl(objOperand2.value))
#End If
                        Case ENUop.opTrace
                            Dim intTraceOperand As Integer
                            Try
                                intTraceOperand = CInt(.Operand)
                            Catch
#If QUICKBASICENGINE_POPCHECK Then
                                interpreter__err_("Trace op has non-integer operand " & _
                                                  _OBJutilities.object2String(.Operand), _
                                                  intIndexIP, _
                                                  objStack, _
                                                  "")
#Else
                                interpreter__err_("Trace op has non-integer operand " & _
                                                  _OBJutilities.object2String(.Operand))
#End If
                                Exit Do
                            End Try
                            With objTraceStack
                                .Pop() : .Push(intTraceOperand)
                            End With
                        Case ENUop.opTracePop
                            With objTraceStack
                                If .Count = 1 Then
#If QUICKBASICENGINE_POPCHECK Then
                                    interpreter__err_("Trace stack underflow", _
                                                      intIndexIP, _
                                                      objStack, _
                                                      "")
#Else
                                    interpreter__err_("Trace stack underflow")
#End If
                                Else
                                    .Pop()
                                End If
                            End With
                        Case ENUop.opTracePush
                            With objTraceStack
                                .Push(.Peek)
                            End With
                        Case ENUop.opTrim
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, Trim(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               Trim(CStr(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case ENUop.opUCase
#If QUICKBASICENGINE_POPCHECK Then
                            interpreter__push_(objStack, UCase(CStr(objFrame(1).value)))
#Else
                            interpreter__push_(objStack, _
                                               intStackTop, _
                                               UCase(CStr(interpreter__pop_(objStack, intStackTop).value)))
#End If
                        Case Else
                            errorHandler_("Compiler error: no support exists for opcode " & _
                                          .opcodeToString, _
                                          "interpreter_", _
                                          "Marking object unusable", _
                                          Nothing)
                            OBJstate.usrState.booUsable = False
                            Exit Do
                    End Select
                Catch
#If QUICKBASICENGINE_POPCHECK Then
                    interpreter__err_("Exception thrown in interpreter: " & _
                                      Err.Number & " " & Err.Description, _
                                      intIndexIP, _
                                      objStack, _
                                      "")
#Else
                    interpreter__err_("Exception thrown in interpreter: " & _
                                      Err.Number & " " & Err.Description)
#End If
                    Return (False)
                End Try
#If QUICKBASICENGINE_POPCHECK Then
                ' --- Dispose of the stack frame
                For intIndex1 = 1 To UBound(objFrame)
                    If (objFrame(intIndex1).Tag Is Nothing) Then
                        disposeVariable_(objFrame(intIndex1))
                    End If
                    objFrame(intIndex1) = Nothing
                Next intIndex1
                ReDim objFrame(0)
#End If
            End With
            intIndexIP += 1 : intCount += 1
        Loop
#If QUICKBASICENGINE_POPCHECK Then
        With objStack
            If .Count = 0 Then Return (True)
            If .Count = 1 Then Return (CType(objStack.Peek, qbVariable.qbVariable))
            Return (False)
        End With
#Else
        If intStackTop = 0 Then Return (True)
        If intStackTop = 1 Then Return objStack(intStackTop)
        Return (False)
#End If
    End Function

    ' -----------------------------------------------------------------
    ' Convert Boolean value to numeric operand on behalf of the
    ' interpreter
    '
    '
    Private Function interpreter__boolean2Operand_(ByVal booOperand As Boolean) As Integer
        Return (CInt(IIf(booOperand, -1, 0)))
    End Function

    ' ----------------------------------------------------------------------
    ' Report errors in the interpreter that are probably due to errors in
    ' the QuickBasic code
    '
    '
    ' --- Polish code, memory and stack available
    Private Overloads Sub interpreter__err_(ByVal strMessage As String, _
                                            ByVal intIndex As Integer, _
                                            ByRef objStack As Stack, _
                                            ByVal strHelp As String)
        raiseEvent_("interpretErrorEvent", strMessage, intIndex, strHelp)
    End Sub
    ' --- Polish code not available
    Private Overloads Sub interpreter__err_(ByVal strMessage As String)
        raiseEvent_("userErrorEvent", strMessage, "")
    End Sub

#If QUICKBASICENGINE_POPCHECK Then
    ' -----------------------------------------------------------------
    ' Allocates or expands the stack frame on behalf of interpreter_
    '
    '
    Private Function interpreter__expandFrame_ _
            (ByRef objFrame() As qbVariable.qbVariable, _
             ByVal intCount As Integer) _
            As Boolean
        Try
            ReDim Preserve objFrame(intCount)
        Catch objException As Exception
            errorHandler_("Cannot allocate or expand the stack frame: " & _
                          Err.Number & " " & Err.Description, _
                          "interpreter__expandFrame_", _
                          "Making object unusable: returning False")
            OBJstate.usrState.booUsable = False : Return (False)
        End Try
        Return (True)
    End Function
#End If

    ' ----------------------------------------------------------------------
    ' Assemble subscripts for an array reference
    '
    '
    Private Function interpreter__getSubscripts_(ByVal objFrame() As qbVariable.qbVariable, _
                                                 ByVal intStartIndex As Integer, _
                                                 ByVal intEndIndex As Integer) _
            As String
        Dim intIndex1 As Integer
        Dim strSubscripts As String
        For intIndex1 = intStartIndex To intEndIndex
            _OBJutilities.append(strSubscripts, _
                                 ",", _
                                 CStr(objFrame(intIndex1).value))
        Next intIndex1
        Return strSubscripts
    End Function

    ' ----------------------------------------------------------------------
    ' Interpreter input
    '
    '
    Private Function interpreter__input_(ByVal objInputQueue As Queue) _
            As qbVariable.qbVariable
        If objInputQueue.Count < 1 Then
            Dim strInput As String
            RaiseEvent interpretInputEvent(Me, strInput)
            If eventLoggingGet_() Then
                extendEventLog_("interpretInputEvent", _
                                "strInput=" & _OBJutilities.string2Display(strInput))
            End If
            If eventLoggingGet_() Then extendEventLog_("interpretInputEvent", _
                                                    "strInput=" & _
                                                    _OBJutilities.enquote(strInput))
            Dim strSplit As String() = Split(strInput, ",")
            If LBound(strSplit) < 0 Then Return (Nothing)
            Dim intIndex1 As Integer
            Dim objNext As qbVariable.qbVariable
            For intIndex1 = LBound(strSplit) To UBound(strSplit)
                objNext = _OBJqbVariable.mkVariable(strSplit(intIndex1))
                objInputQueue.Enqueue(objNext)
            Next intIndex1
        End If
        With objInputQueue
            If .Count < 1 Then Return (Nothing)
            Return (CType(.Dequeue, qbVariable.qbVariable))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Retrieve source code associated with a Polish operation
    '
    '
    Private Function interpreter__polish2Source_(ByVal intIndex As Integer) _
            As String
        With OBJstate.usrState
            Dim objHandle As qbPolish.qbPolish = _
                CType(.colPolish.Item(intIndex), qbPolish.qbPolish)
            With objHandle
                Return (OBJstate.usrState.objScanner.sourceMid(.TokenStartIndex, .TokenLength))
            End With
        End With
    End Function

#If QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Pop the interpreter stack when PopCheck is in effect 
    '
    '
    Private Function interpreter__pop_(ByRef objStack As Stack, _
                                       ByVal strTemplate As String, _
                                       ByRef objFrame() As qbVariable.qbVariable) _
            As Boolean
        If strTemplate = "" Then Return (True)
        Dim strTemplateSplit() As String
        Try
            strTemplateSplit = Split(strTemplate, ",")
        Catch objException As Exception
            errorHandler_("Cannot split template", _
                        "interpreter__pop_", _
                        "Marking object unusable: returning Nothing", _
                        objException)
            OBJstate.usrState.booUsable = False
            Return (False)
        End Try
        Dim intIndex1 As Integer
        If InStr("," & strTemplate & ",", ",a,") = 0 Then
            ' Fast pop
            If Not interpreter__expandFrame_(objFrame, UBound(strTemplateSplit) + 1) Then
                Return (False)
            End If
            For intIndex1 = UBound(objFrame) To 1 Step -1
                objFrame(intIndex1) = CType(objStack.Pop, qbVariable.qbVariable)
            Next intIndex1
            Return True
        End If
        If Not interpreter__expandFrame_(objFrame, 0) Then
            Return (False)
        End If
        Dim booOK As Boolean
        Dim strNext As String
        Dim strTemplateDesc As String
        For intIndex1 = UBound(strTemplateSplit) To 0 Step -1
            strNext = LCase(Trim(strTemplateSplit(intIndex1)))
            If strNext = "a" Then
                booOK = interpreter__pop__arraySubscripts_(objStack, objFrame)
            Else
                If Not interpreter__expandFrame_(objFrame, UBound(objFrame) + 1) Then
                    Return False
                End If
                Try
                    objFrame(UBound(objFrame)) = CType(objStack.Pop, qbVariable.qbVariable)
                Catch objException As Exception
                    errorHandler_("Cannot pop interpreter stack: " & Err.Number & " " & Err.Description, _
                                "interpreter__pop_", _
                                "Marking object unusable: returning Nothing", _
                                objException)
                    OBJstate.usrState.booUsable = False : Return (False)
                End Try
                Select Case strNext
                    Case "x"
                        strTemplateDesc = "anything"
                        booOK = True
                    Case "s"
                        booOK = objFrame(UBound(objFrame)).isScalar
                        strTemplateDesc = "scalar"
                    Case "n"
                        booOK = objFrame(UBound(objFrame)).isaNumber
                        strTemplateDesc = "number"
                    Case "i"
                        booOK = objFrame(UBound(objFrame)).isAnUnsignedInteger
                        strTemplateDesc = "unsigned integer"
                    Case Else
                        booOK = (LCase(objFrame(UBound(objFrame)).Dope.VariableType.ToString) = strNext)
                End Select
            End If
            If Not booOK Then
                errorHandler_("Owing to a compiler error, the stack frame " & _
                              interpreter__pop__stackFrame2String_(objFrame) & " " & _
                              "contains a bad entry at position " & intIndex1 & ": " & _
                              "expected " & _
                              strTemplateDesc & ": " & _
                              "found " & _
                              objFrame(intIndex1).ToString, _
                              "", _
                              "Marking object unusable: returning False", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return (False)
            End If
        Next intIndex1
        Dim intIndex2 As Integer = UBound(objFrame)
        Dim objExchange As qbVariable.qbVariable
        For intIndex1 = 1 To UBound(objFrame)
            If intIndex1 >= intIndex2 Then Exit For
            objExchange = objFrame(intIndex1)
            objFrame(intIndex1) = objFrame(intIndex2)
            objFrame(intIndex2) = objExchange
        Next intIndex1
        Return (True)
    End Function
#End If

#If Not QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Pop the interpreter stack when PopCheck is not in effect 
    '
    '
    ' --- Pop one entry and return it as the function value
    Private Overloads Function interpreter__pop_(ByRef objStack() As qbVariable.qbVariable, _
                                                 ByRef intStackTop As Integer) _
            As qbVariable.qbVariable
        Return interpreter__pop_(objStack, intStackTop, 1)
    End Function
    ' --- Pop multiple entries and return the last
    Private Overloads Function interpreter__pop_(ByRef objStack() As qbVariable.qbVariable, _
                                                 ByRef intStackTop As Integer, _
                                                 ByVal intCount As Integer) _
            As qbVariable.qbVariable
        If intStackTop < intCount Then
            interpreter__err_("Stack underflow") : Return Nothing
        End If
        Dim objReturn As qbVariable.qbVariable = objStack(intStackTop - intCount + 1)
        intStackTop -= intCount
        If UBound(objStack) - intStackTop > INTERPRETER_STACK_MAXFREE Then
            Try
                ReDim Preserve objStack(UBound(objStack) + INTERPRETER_STACK_BLOCK)
            Catch
                interpreter__err_("Cannot trim stack: " & Err.Number & " " & Err.Description)
            End Try
        End If
        Return objReturn
    End Function
#End If

#If QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Pop an array subscript list
    '
    '
    ' Expects I(1), I(2)...I(n),count,qbVariable from bottom to top
    '
    '
    Private Function interpreter__pop__arraySubscripts_(ByRef objStack As Stack, _
                                                        ByRef objFrame() As qbVariable.qbVariable) _
            As Boolean
        With objStack
            ' --- Pops the qbVariable
            Dim objArray As qbVariable.qbVariable
            Try
                objArray = CType(objStack.Pop, qbVariable.qbVariable)
            Catch objException As Exception
                errorHandler_("Cannot pop top of stack", _
                              "interpreter__pop__arraySubscripts_", _
                              "Making object unusable: returning False", _
                              objException)
                OBJstate.usrState.booUsable = False : Return (False)
            End Try
            ' --- Pops the count
            Dim intCount As Integer
            Try
                intCount = CInt(CType(.Pop, qbVariable.qbVariable).value)
            Catch objException As Exception
                errorHandler_("Compiler error: array subscript stack frame is " & _
                              "invalid: no count found: " & _
                              Err.Number & " " & Err.Description, _
                              "interpreter__pop__arraySubscripts_", _
                              "Making object unusable: returning False", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return (False)
            End Try
            If intCount > .Count Then
                errorHandler_("Compiler error: array subscript stack frame is " & _
                              "invalid: stack underflow", _
                              "interpreter__pop__arraySubscripts_", _
                              "Making object unusable: returning False", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return (False)
            End If
            If Not objArray.Dope.isArray _
               OrElse _
               objArray.Dope.Dimensions <> intCount Then
                errorHandler_("Compiler error: array subscript stack frame is " & _
                              "invalid: nonarray, or wrong number of subscripts", _
                              "interpreter__pop__arraySubscripts_", _
                              "Making object unusable: returning False", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return (False)
            End If
            ' --- Create the stack frame as I(1), I(2)...I(n),qbVariable
            If Not interpreter__expandFrame_(objFrame, intCount + 1) Then Return (False)
            objFrame(UBound(objFrame)) = objArray
            Dim intIndex1 As Integer
            For intIndex1 = UBound(objFrame) - 1 To 1 Step -1
                Try
                    objFrame(intIndex1) = CType(.Pop, qbVariable.qbVariable)
                Catch objException As Exception
                    errorHandler_("Cannot pop subscript: " & _
                                  Err.Number & " " & Err.Description, _
                                  "interpreter__pop__arraySubscripts_", _
                                  "Making object unusable: returning False", _
                                  objException)
                    OBJstate.usrState.booUsable = False : Return (False)
                End Try
            Next intIndex1
            Return (True)
        End With
    End Function
#End If

#If QUICKBASICENGINE_POPCHECK Then
    ' -----------------------------------------------------------------
    ' Convert a stack frame to a string on behalf of the interpreter
    '
    '
    Private Function interpreter__pop__stackFrame2String_(ByVal objFrame() As qbVariable.qbVariable) _
            As String
        Dim intIndex1 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To UBound(objFrame)
            strOutstring = _OBJutilities.append(strOutstring, _
                                                ",", _
                                                objFrame(intIndex1).ToString)
        Next intIndex1
        Return (strOutstring)
    End Function
#End If

#If QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Push a value on the stack (pop check in effect)
    '
    '
    ' The value is pushed on the stack in one of three ways:
    '
    '
    '      *  If the value is a qbVariable by default it is pushed "by
    '         value": it is here cloned and pushed
    '
    '      *  If the booByVal option is present and False the variable
    '         is not cloned before it is pushed "by reference"
    '
    '      *  If the value is not a qbVariable it is converted to a
    '         qbVariable, and pushed by value
    '
    '
    ' In all cases, the variable's Tag property will indicate whether it is
    ' a value or a reference:
    '
    '
    '      *  If the variable is a cloned value then its Tag will be Nothing
    '
    '      *  Otherwise the Tag will be the Tag of the reference, and, this
    '         will be the storage address of the variable
    '
    '
    ' --- Push value 
    Private Overloads Sub interpreter__push_(ByRef objStack As Stack, _
                                             ByVal objValue As Object)
        interpreter__push_(objStack, objValue, True)
    End Sub
    ' --- Exposes option to push by reference 
    Private Overloads Sub interpreter__push_(ByRef objStack As Stack, _
                                             ByVal objValue As Object, _
                                             ByVal booByVal As Boolean)
        Dim objHandle As qbVariable.qbVariable
        If (TypeOf objValue Is qbVariable.qbVariable AndAlso booByVal) Then
            ' Copy and push the copied value 
            Try
                objHandle = CType(objValue, qbVariable.qbVariable).clone
            Catch objException As Exception
                errorHandler_("Cannot clone qbVariable: " & _
                              Err.Number & " " & Err.Description, _
                              "interpreter__push_", _
                              "Marking object unusable", _
                              objException)
                OBJstate.usrState.booUsable = False : Return
            End Try
            objHandle.Tag = Nothing
        ElseIf (TypeOf objValue Is qbVariable.qbVariable) Then
            objHandle = CType(objValue, qbVariable.qbVariable)
            If (TypeOf objHandle.Tag Is System.Boolean) AndAlso CBool(objHandle.Tag) Then
                errorHandler_("Cannot push qbVariable by reference because " & _
                              "its Tag is Boolean and True", _
                              "interpreter__push_", _
                              "Marking object unusable", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return
            End If
        Else
            objHandle = _OBJqbVariable.mkVariableFromValue(objValue)
        End If
        Try
            objStack.Push(objHandle)
        Catch objException As Exception
            errorHandler_("Interpreter can't push its stack: " & _
                          Err.Number & " " & Err.Description, _
                          "interpreter__push_", _
                          "Marking object unusable", _
                          objException)
        End Try
    End Sub
#End If

#If Not QUICKBASICENGINE_POPCHECK Then
    ' ----------------------------------------------------------------------
    ' Push a value on the stack (pop check not in effect)
    '
    '
    ' The value is pushed on the stack in one of three ways:
    '
    '
    '      *  If the value is a qbVariable by default it is pushed "by
    '         value": it is here cloned and pushed
    '
    '      *  If the booByVal option is present and False the variable
    '         is not cloned before it is pushed "by reference"
    '
    '      *  If the value is not a qbVariable it is converted to a
    '         qbVariable, and pushed by value
    '
    '
    ' In all cases, the variable's Tag property will indicate whether it is
    ' a value or a reference:
    '
    '
    '      *  If the variable is a cloned value then its Tag will be Nothing
    '
    '      *  Otherwise the Tag will be the Tag of the reference, and, this
    '         will be the storage address of the variable
    '
    '
    ' --- Push value 
    Private Overloads Sub interpreter__push_(ByRef objStack() As qbVariable.qbVariable, _
                                             ByRef intStackTop As Integer, _
                                             ByVal objValue As Object)
        interpreter__push_(objStack, intStackTop, objValue, True)
    End Sub
    ' --- Exposes option to push by reference 
    Private Overloads Sub interpreter__push_(ByRef objStack() As qbVariable.qbVariable, _
                                             ByRef intStackTop As Integer, _
                                             ByVal objValue As Object, _
                                             ByVal booByVal As Boolean)
        intStackTop += 1
        If UBound(objStack) <= intStackTop Then
            Try
                ReDim Preserve objStack(UBound(objStack) + INTERPRETER_STACK_BLOCK)
            Catch objException As Exception
                errorHandler_("Cannot resize the stack: " & _
                              Err.Number & " " & Err.Description, _
                              "interpreter__push_", _
                              "Marking object unusable", _
                              objException)
                OBJstate.usrState.booUsable = False : Return
            End Try
        End If
        Dim objHandle As qbVariable.qbVariable
        If (TypeOf objValue Is qbVariable.qbVariable AndAlso booByVal) Then
            ' Copy and push the copied value 
            Try
                objHandle = CType(objValue, qbVariable.qbVariable).clone
            Catch objException As Exception
                errorHandler_("Cannot clone qbVariable: " & _
                              Err.Number & " " & Err.Description, _
                              "interpreter__push_", _
                              "Marking object unusable", _
                              objException)
                OBJstate.usrState.booUsable = False : Return
            End Try
            objHandle.Tag = Nothing
        ElseIf (TypeOf objValue Is qbVariable.qbVariable) Then
            objHandle = CType(objValue, qbVariable.qbVariable)
            If (TypeOf objHandle.Tag Is System.Boolean) AndAlso CBool(objHandle.Tag) Then
                errorHandler_("Cannot push qbVariable by reference because " & _
                              "its Tag is Boolean and True", _
                              "interpreter__push_", _
                              "Marking object unusable", _
                              Nothing)
                OBJstate.usrState.booUsable = False : Return
            End If
        Else
            objHandle = _OBJqbVariable.mkVariableFromValue(objValue)
        End If
        objStack(intStackTop) = objHandle
    End Sub
#End If

    ' ----------------------------------------------------------------------
    ' Type resolution for compare
    '
    '
    ' If both operands 1 and 2 are numbers, this function returns True:
    ' if one or both are strings, this function converts both to strings,
    ' and returns False.
    '
    '
    Private Function interpreter__resolveTypes_(ByVal objOperand1 As qbVariable.qbVariable, _
                                                ByVal objOperand2 As qbVariable.qbVariable) _
            As Boolean
        If objOperand1.isaNumber AndAlso objOperand2.isaNumber Then Return (True)
        With objOperand1
            .fromString(_OBJutilities.enquote(CStr(.value)))
        End With
        With objOperand2
            .fromString(_OBJutilities.enquote(CStr(.value)))
        End With
        Return (False)
    End Function

#If Not QUICKBASICENGINE_EXTENSION Then
    ' ----------------------------------------------------------------------
    ' Truncate string to 64KB
    '
    '
    Private Function interpreter__stringTruncate_(ByRef objStringHandle As qbVariable.qbVariable) _
            As Boolean
        With objStringHandle
            .valueSet(Mid(CStr(.value), 1, 65535))
            Return True
        End With
    End Function
#End If

#End Region

    ' ----------------------------------------------------------------------
    ' Loop event interface: check for the stopped state: return True if
    ' object has been stopped
    '
    '
    Public Function loopEventInterface_(ByVal strActivity As String, _
                                        ByVal strEntity As String, _
                                        ByVal intNumber As Integer, _
                                        ByVal intCount As Integer, _
                                        ByVal intLevel As Integer, _
                                        ByVal strComment As String) As Boolean
        Dim booStopping As Boolean = (Me.getThreadStatus = "Stopping")
        Dim strCommentWork As String = strComment
        If booStopping Then
            strCommentWork = strComment & _
                             CStr(IIf(strComment = "", "", vbNewLine & vbNewLine)) & _
                             "Note that this thread is stopping"
        End If
        raiseEvent_("loopEvent", _
                    strActivity, _
                    strEntity, _
                    intNumber, _
                    intCount, _
                    intLevel, _
                    strCommentWork)
        Return (booStopping)
    End Function

    ' ----------------------------------------------------------------------
    ' Make the collection if it does not exist
    '
    '
    Private Function mkCollection_(ByRef colCollection As Collection, _
                                   ByVal strName As String) As Boolean
        Try
            colCollection = New Collection
        Catch objException As Exception
            errorHandler_("Cannot create the " & strName & " collection: " & _
                            Err.Number & " " & Err.Description, _
                            "mkCollection_", _
                            "Marking quickBasicEngine not usable and returning False", _
                            objException)
            Return (False)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Make the event log
    '
    '
    Private Function mkEventLog_() As Boolean
        Try
            OBJstate.usrState.colEventLog = New Collection
            Return (True)
        Catch objException As Exception
            errorHandler_("Cannot create an event log: " & _
                            Err.Number & " " & Err.Description, _
                            "EventLogging Set", _
                            "Continuing with no event logging", _
                            objException)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Make the scanner object  
    '
    '
    Private Function mkScanner_() As Boolean
        With OBJstate.usrState
            Try
                OBJscanner = New qbScanner.qbScanner
                .objScanner = OBJscanner
            Catch objException As Exception
                errorHandler_("Cannot create scanner: " & _
                              Err.Number & " " & Err.Description, _
                              "mkScanner_", _
                              "Marking quickBasicEngine not usable and returning False", _
                              objException)
                Return (False)
            End Try
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Run some code in MSIL  
    '
    '
    Private Function msilRun_() As Object
        With OBJstate.usrState
            If (.colPolish Is Nothing) Then
                errorHandler_("Cannot run MSIL code: no Polish code is available", _
                              "msilRun_", _
                              "Returning Nothing", _
                              Nothing)
                Return Nothing
            End If
            Dim objAsmName As AssemblyName
            Dim objAsm As AssemblyBuilder
            Dim objClass As TypeBuilder
            Dim objILgenerator As ILGenerator
            Dim objMethod As MethodBuilder
            Dim objModule As ModuleBuilder
            Try
                objAsmName = New AssemblyName
                objAsmName.Name = "msilRun"
                objAsmName.Version = New Version("1.0.0.0")
                objAsm = AppDomain.CurrentDomain.DefineDynamicAssembly(objAsmName, _
                                                                       AssemblyBuilderAccess.Run)
                objModule = objAsm.DefineDynamicModule(objAsmName.Name)
                objClass = objModule.DefineType(objAsmName.Name, TypeAttributes.Public)
                objMethod = objClass.DefineMethod(objAsmName.Name & "_", _
                                                  MethodAttributes.Public, _
                                                  Type.GetType("System.Double"), _
                                                  Nothing)
                objILgenerator = objMethod.GetILGenerator
            Catch objException As Exception
                errorHandler_("Not able to initialize MSIL generation: " & _
                              Err.Number & " " & Err.Description, _
                              "msilRun_", _
                              "Returning nothing", _
                              objException)
            End Try
            Dim intIndex1 As Integer
            Dim objArgument As Object
            Dim objNextOpcode As OpCode
            Dim objNextValue As Object
            With .colPolish
                For intIndex1 = 1 To .Count
                    loopEventInterface_("Generating MSIL code", _
                                        "collection item", _
                                        intIndex1, _
                                        .Count, _
                                        0, _
                                        "")
                    With CType(.Item(intIndex1), qbPolish.qbPolish)
                        If msilRun__qbOpcode2MSIL_(.Opcode, objNextOpcode) Then
                            objILgenerator.Emit(objNextOpcode)
                        Else
                            If .Opcode = ENUop.opPushLiteral Then
                                If UCase(.Operand.GetType.ToString) = "QBVARIABLE.QBVARIABLE" Then
                                    objNextValue = CType(.Operand, qbVariable.qbVariable).value
                                Else
                                    objNextValue = .Operand
                                End If
                                objILgenerator.Emit(OpCodes.Ldc_R8, CDbl(objNextValue))
                            End If
                        End If
                    End With
                Next intIndex1
                If intIndex1 <= .Count Then
                    errorHandler_("Not able to convert Polish code to MSIL", _
                                  "msilRun_", _
                                  "Returning Nothing", _
                                  Nothing)
                    Return Nothing
                End If
                objILgenerator.Emit(OpCodes.Ret)
                Dim objReturn As Object
                Try
                    objClass.CreateType()
                    Dim objType As Type = objAsm.GetType(objClass.Name)
                    Dim objInstance As Object = Activator.CreateInstance(objType)
                    Dim objMethodInfo As MethodInfo = objType.GetMethod(objMethod.Name)
                    objReturn = objMethodInfo.Invoke(objInstance, Nothing)
                Catch objException As Exception
                    errorHandler_("Not able to run MSIL: " & _
                                  Err.Number & " " & Err.Description, _
                                  "msilRun_", _
                                  "Returning Nothing", _
                                  Objexception)
                    Return Nothing
                End Try
                Return (objReturn)
            End With
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Convert a set of operands, small at this writing, to MSIL  
    '
    '
    ' Note that at this writing, END is translated to NOP.
    '
    '
    Private Function msilRun__qbOpcode2MSIL_(ByVal enuPolishOpcode As qbOp.qbOp.ENUop, _
                                             ByRef enuMSILopcode As OpCode) As Boolean
        Select Case enuPolishOpcode
            Case ENUop.opAdd : enuMSILopcode = OpCodes.Add
            Case ENUop.opAnd : enuMSILopcode = OpCodes.And
            Case ENUop.opDivide : enuMSILopcode = OpCodes.Div
            Case ENUop.opEnd : enuMSILopcode = OpCodes.Nop
            Case ENUop.opMultiply : enuMSILopcode = OpCodes.Mul
            Case ENUop.opNegate : enuMSILopcode = OpCodes.Neg
            Case ENUop.opNot : enuMSILopcode = OpCodes.Not
            Case ENUop.opOr : enuMSILopcode = OpCodes.Or
            Case ENUop.opSubtract : enuMSILopcode = OpCodes.Sub
            Case Else
                Return False
        End Select
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Returns a usable collection key for a string or a number
    '
    '
    Private Function object2Key_(ByVal objValue As Object) As String
        Try
            Return ("_" & CStr(objValue))
        Catch objException As Exception
            errorHandler_("Internal compiler error: " & _
                          "object cannot be converted to a string", _
                          "", _
                          "Marking object unusable: returning a null string", _
                          objException)
            OBJstate.usrState.booUsable = False
            Return ("")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Convert the object's state to eXtended Markup Language
    '
    '
    Private Function object2XML_(ByVal booAboutComment As Boolean, _
                                 ByVal booStateComment As Boolean) _
           As String
        SyncLock OBJstate
            If Not checkUsable_("object2XML", "Returning XML comment") Then
                Return _OBJutilities.mkXMLComment("Unusable object instance")
            End If
            raiseEvent_("msgEvent", "Converting object to XML")
            With OBJstate.usrState
                Dim strGenerateNOPs As String = "True"
#If QUICKBASICENGINE_EXTENSION Then
                strGenerateNOPs = CStr(.booGenerateNOPs)
#End If
                Dim strScannerXML As String = "Nothing"
                If Not (.objScanner Is Nothing) Then
                    strScannerXML = Replace(.objScanner.object2XML, _
                                            vbNewLine, _
                                            vbNewLine & "    ")
                End If
                Return (_OBJutilities.objectInfo2XML(ClassName, _
                                                        ClassName & vbNewLine & vbNewLine & _
                                                        _OBJutilities.soft2HardParagraph(About, 60), _
                                                        booAboutComment, booStateComment, _
                                                        "booUsable", _
                                                        "Indicates object usability", _
                                                        CStr(.booUsable), _
                                                        "strName", _
                                                        "Object instance's name", _
                                                        .strName, _
                                                        "objScanner", _
                                                        "Scanner object", _
                                                        vbNewLine & _
                                                        strScannerXML & _
                                                        vbNewLine, _
                                                        "colPolish", _
                                                        "Polish collection", _
                                                        vbNewLine & _
                                                        Replace(object2XML_collection2XML_(OBJstate.usrState.colPolish), _
                                                                vbNewLine, _
                                                                vbNewLine & "    ") & _
                                                        vbNewLine, _
                                                        "colVariables", _
                                                        "Variables", _
                                                        vbNewLine & _
                                                        Replace(object2XML_collection2XML_(OBJstate.usrState.colVariables), _
                                                                vbNewLine, _
                                                                vbNewLine & "    ") & _
                                                        vbNewLine, _
                                                        "objCollectionUtilities", _
                                                        "Collection utilities", _
                                                        _OBJutilities.ellipsis(.objCollectionUtilities.Name, 16), _
                                                        "booCompiled", _
                                                        "Indicates compiled status", _
                                                        CStr(.booCompiled), _
                                                        "booAssembled", _
                                                        "Indicates assembler status", _
                                                        CStr(.booAssembled), _
                                                        "enuSourceCodeType", _
                                                        "Source code type", _
                                                        .enuSourceCodeType.ToString, _
                                                        "booConstantFolding", _
                                                        "Constant folding", _
                                                        CStr(.booConstantFolding), _
                                                        "booDegenerateOpRemoval", _
                                                        "Removal of ""degenerate"" operations", _
                                                        CStr(.booDegenerateOpRemoval), _
                                                        "objImmediateResult", _
                                                        "Value of recent immediate expression", _
                                                        _OBJutilities.object2String(.objImmediateResult), _
                                                        "booExplicit", _
                                                        "Indicates whether Option Explicit is in effect", _
                                                        CStr(.booExplicit), _
                                                        "usrSubFunction", _
                                                        "Subroutine/function table", _
                                                        object2XML_subFunctionTable_(.usrSubFunction), _
                                                        "colSubFunctionIndex", _
                                                        "Subroutine/function table index", _
                                                        .objCollectionUtilities.collection2String(.colSubFunctionIndex, _
                                                                                                  booReadable:=True), _
                                                        "usrConstantExpression", _
                                                        "Constant expression table", _
                                                        object2XML_constantExpressionTable_(.usrConstantExpression), _
                                                        "colConstantExpressionIndex", _
                                                        "Constant expression table index", _
                                                        .objCollectionUtilities.collection2String(.colConstantExpressionIndex, _
                                                                                                  booReadable:=True), _
                                                        "intOldestConstantIndex", _
                                                        "Index of the oldest constant expression", _
                                                        CStr(.intOldestConstantIndex), _
                                                        "intLabelSeq", _
                                                        "Label sequence number", _
                                                        CStr(.intLabelSeq), _
                                                        "colReadData", _
                                                        "Read Data queue", _
                                                        .objCollectionUtilities.collection2String(.colReadData, _
                                                                                                  booReadable:=True), _
                                                        "intReadDataIndex", _
                                                        "Next Read Data index", _
                                                        CStr(.intReadDataIndex), _
                                                        "colLabel", _
                                                        "Label collection for CoGo", _
                                                        .objCollectionUtilities.collection2String(.colLabel, _
                                                                                                  booReadable:=True), _
                                                        "booAssemblyRemovesCode", _
                                                        "True: use the assembler to remove comments and labels", _
                                                        CStr(.booAssemblyRemovesCode), _
                                                        "colEventLog", _
                                                        "Event log", _
                                                        .objCollectionUtilities.collection2String(.colEventLog, _
                                                                                                  booReadable:=True), _
                                                        "booInspectCompilerObjects", _
                                                        "True: inspect compiler objects", _
                                                        CStr(.booInspectCompilerObjects), _
                                                        "threadStatus", _
                                                        "Thread status", _
                                                        OBJthreadStatus.ToString, _
                                                        "generateNOPs", _
                                                        "True: generate NOPs and REMs", _
                                                        strGenerateNOPs))
            End With
        End SyncLock
    End Function

    ' -----------------------------------------------------------------
    ' Polish collection or variable to eXtended Markup Language on 
    ' behalf of object2XML
    '
    '
    Private Function object2XML_collection2XML_(ByVal colCollection As Collection) _
            As String
        If (colCollection Is Nothing) Then Return ("Nothing")
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch objException As Exception
            errorHandler_("Cannot make String Builder: " & _
                          Err.Number & " " & Err.Description, _
                          "object2XML_polish2XML_", _
                          "Returning XML comment", _
                          objException)
            Return _OBJutilities.mkXMLComment("Error in creating")
        End Try
        With colCollection
            If .Count > 0 Then
                Dim booPolish As Boolean
                Dim objHandlePolish As qbPolish.qbPolish
                Dim objHandleVariable As qbVariable.qbVariable
                Dim strName As String
                If (TypeOf .Item(1) Is qbPolish.qbPolish) Then
                    objHandlePolish = CType(.Item(1), qbPolish.qbPolish)
                    strName = objHandlePolish.ClassName
                    booPolish = True
                ElseIf (TypeOf .Item(1) Is qbVariable.qbVariable) Then
                    objHandleVariable = CType(.Item(1), qbVariable.qbVariable)
                    strName = objHandleVariable.ClassName
                    booPolish = False
                Else
                    errorHandler_("Programming error: unsupported collection member type " & _
                                  "occurs at the start of a collection", _
                                  "object2XML_collection2XML_", _
                                  "Marking object not usable", _
                                  Nothing)
                    OBJstate.usrState.booUsable = False
                    Return (_OBJutilities.mkXMLComment("Internal error"))
                End If
                Dim intAlignLength As Integer = CInt(Math.Log10(.Count) + 1)
                Dim intIndex1 As Integer
                Dim strXML As String
                For intIndex1 = 1 To .Count
                    raiseEvent_("loopEvent", _
                                "Converting collection to XML", _
                                "item", _
                                intIndex1, _
                                .Count, _
                                0, _
                                "")
                    If booPolish Then
                        objHandlePolish = CType(.Item(intIndex1), qbPolish.qbPolish)
                        strXML = objHandlePolish.object2XML(False, True)
                    Else
                        objHandleVariable = CType(.Item(intIndex1), _
                                                  qbVariable.qbVariable)
                        strXML = objHandleVariable.object2XML
                    End If
                    _OBJutilities.append(objStringBuilder, _
                                         vbNewLine, _
                                        _OBJutilities.mkXMLElement _
                                        (strName & _
                                         _OBJutilities.alignRight _
                                         (CStr(intIndex1), intAlignLength, "0"), _
                                         strXML))
                Next intIndex1
            End If
            Return (objStringBuilder.ToString)
        End With
    End Function

    ' ------------------------------------------------------------------
    ' Convert constant expression table to XML
    '
    '
    Private Function object2XML_constantExpressionTable_(ByVal usrTable() As TYPconstantExpression) _
            As String
        Dim intIndex1 As Integer
        Dim intUBound As Integer = -1
        Try
            intUBound = UBound(usrTable)
        Catch : End Try
        If intUBound = -1 Then Return ("Unallocated")
        Dim strOutstring As String
        For intIndex1 = 1 To intUBound
            raiseEvent_("loopEvent", _
                        "Converting constant expression table to XML", _
                        "entry", _
                        intIndex1, _
                        intUBound, _
                        0, _
                        "")
            With usrTable(intIndex1)
                _OBJutilities.append(strOutstring, vbNewLine, _
                                     _OBJutilities.objectInfo2XML _
                                     ("constantExpression" & _
                                      _OBJutilities.alignRight(CStr(intIndex1), 4, "0"), _
                                      "", _
                                      False, False, _
                                      "strConstantExpression", _
                                      "", _
                                      .strConstantExpression, _
                                      "objValue", _
                                      "", _
                                      _OBJutilities.object2String(.objValue.value)))
            End With
        Next intIndex1
        Return (_OBJutilities.mkXMLElement("constantExpressionTable", _
                                           strOutstring))
    End Function

    ' ------------------------------------------------------------------
    ' Convert the formal parameter table to XML
    '
    '
    Private Function object2XML_formalParms_(ByVal usrTable() As TYPformalParameters) _
            As String
        Dim intIndex1 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To UBound(usrTable)
            With usrTable(intIndex1)
                _OBJutilities.append(strOutstring, vbNewLine, _
                                     _OBJutilities.objectInfo2XML _
                                     ("formalParameter" & _
                                      _OBJutilities.alignRight(CStr(intIndex1), 4, "0"), _
                                      "booByVal: indicates whether formal parameter is called by value" & _
                                      vbNewLine & _
                                      "strName: Formal parameter's name", _
                                      True, _
                                      False, _
                                      "booByVal", _
                                      "", _
                                      CStr(.booByVal), _
                                      "strName", _
                                      "", _
                                      .strName))
            End With
        Next intIndex1
        Return (_OBJutilities.mkXMLElement("subFunctionTable", _
                                          strOutstring, _
                                          72))
    End Function

    ' ------------------------------------------------------------------
    ' Convert the subroutine and function table to XML
    '
    '
    Private Function object2XML_subFunctionTable_(ByVal usrTable() As TYPsubFunction) _
            As String
        Dim intIndex1 As Integer
        Dim intUBound As Integer = -1
        Try
            intUBound = UBound(usrTable)
        Catch : End Try
        If intUBound = -1 Then Return ("Unallocated")
        Dim strOutstring As String
        For intIndex1 = 1 To intUBound
            With usrTable(intIndex1)
                _OBJutilities.append(strOutstring, vbNewLine, _
                                     _OBJutilities.objectInfo2XML _
                                     ("subFunction" & _
                                        _OBJutilities.alignRight(CStr(intIndex1), 4, "0"), _
                                        "booFunction: True for function: False for subroutine" & _
                                        vbNewLine & _
                                        "strName: Subroutine/function name" & _
                                        vbNewLine & _
                                        "usrFormalParameters: Formal parameters" & _
                                        vbNewLine & _
                                        "intLocation: Location", _
                                        True, False, _
                                        "booFunction", _
                                        "", _
                                        CStr(.booFunction), _
                                        "strName", _
                                        "", _
                                        .strName, _
                                        "usrFormalParameters", _
                                        "", _
                                        object2XML_formalParms_ _
                                        (.usrFormalParameters), _
                                        "intLocation", _
                                        "", _
                                        CStr(.intLocation)))
            End With
        Next intIndex1
        Return (_OBJutilities.mkXMLElement("subFunctionTable", _
                                          strOutstring, _
                                          72))
    End Function

    ' ----------------------------------------------------------------------
    ' RaiseEvent interface
    '
    '
    Private Sub raiseEvent_(ByVal strName As String, _
                            ByVal ParamArray objOperands() As Object)
        Dim booBoolean1 As Boolean
        Dim intInteger1 As Integer
        Dim intInteger2 As Integer
        Dim intInteger3 As Integer
        Dim intInteger4 As Integer
        Dim intInteger5 As Integer
        Dim intInteger6 As Integer
        Dim intInteger7 As Integer
        Dim strOperands As String
        Dim strString1 As String
        Dim strString2 As String
        Dim strString3 As String
        Select Case UCase(strName)
            Case "CODECHANGEEVENT"
                Dim objOp As qbPolish.qbPolish = CType(objOperands(0), qbPolish.qbPolish)
                intInteger1 = CInt(objOperands(1))
                RaiseEvent codeChangeEvent(Me, objOp, intInteger1)
                If eventLoggingGet_() Then
                    strOperands = "objOp=" & _
                                    _OBJutilities.enquote(objOp.ToString) & _
                                    vbNewLine & _
                                    "objOpIndex=" & _
                                    objOperands(1).ToString
                End If
            Case "CODEGENEVENT"
                Dim objOp As qbPolish.qbPolish = CType(objOperands(0), qbPolish.qbPolish)
                RaiseEvent codeGenEvent(Me, objOp)
                If eventLoggingGet_() Then
                    strOperands = "objOp=" & _OBJutilities.enquote(objOp.ToString)
                End If
            Case "CODEREMOVEEVENT"
                intInteger1 = CInt(objOperands(0))
                RaiseEvent codeRemoveEvent(Me, intInteger1)
                If eventLoggingGet_() Then
                    strOperands = "objOpIndex=" & intInteger1.ToString
                End If
            Case "COMPILEERROREVENT"
                strString1 = objOperands(0).ToString
                intInteger1 = CInt(objOperands(1))
                intInteger2 = CInt(objOperands(2))
                intInteger3 = CInt(objOperands(3))
                strString2 = objOperands(4).ToString
                strString3 = objOperands(5).ToString
                RaiseEvent compileErrorEvent(Me, _
                                             strString1, _
                                             intInteger1, _
                                             intInteger2, _
                                             intInteger3, _
                                             strString2, _
                                             strString3)
                If eventLoggingGet_() Then
                    strOperands = "strMessage=" & _OBJutilities.enquote(strString1) & vbNewLine & _
                                    "intIndex=" & intInteger1 & vbNewLine & _
                                    "intContextLength=" & intInteger2 & vbNewLine & _
                                    "intLinenumber=" & intInteger3 & vbNewLine & _
                                    "strHelp=" & _OBJutilities.enquote(strString2) & vbNewLine & _
                                    "strCode=" & _OBJutilities.enquote(strString3)
                End If
            Case "INTERPRETCLSEVENT"
                RaiseEvent interpretClsEvent(Me)
            Case "INTERPRETERROREVENT"
                strString1 = objOperands(0).ToString
                intInteger1 = CInt(objOperands(1))
                strString2 = objOperands(2).ToString
                RaiseEvent interpretErrorEvent(Me, strString1, intInteger1, strString2)
                If eventLoggingGet_() Then
                    strOperands = "strMessage=" & _OBJutilities.enquote(strString1) & vbNewLine & _
                                    "intIndex=" & intInteger1 & vbNewLine & _
                                    "strHelp=" & _OBJutilities.enquote(strString2)
                End If
            Case "INTERPRETPRINTEVENT"
                strString1 = objOperands(0).ToString
                RaiseEvent interpretPrintEvent(Me, strString1)
                If eventLoggingGet_() Then
                    strOperands = "strOutstring=" & _OBJutilities.enquote(strString1)
                End If
#If QUICKBASICENGINE_POPCHECK Then
            Case "INTERPRETTRACEEVENT"
                intInteger1 = CInt(objOperands(0))
                Dim objStack As Stack = CType(objOperands(1), Stack)
                Dim colCollection As Collection = CType(objOperands(2), Collection)
                RaiseEvent interpretTraceEvent(Me, _
                                               intInteger1, _
                                               objStack, _
                                               colCollection)
                If eventLoggingGet_() Then
                    strOperands = "intIndex=" & intInteger1 & vbNewLine & _
                                    "objStack=" & _
                                    Replace(stack2String(objStack), vbNewLine, "\") & _
                                    vbNewLine & _
                                    "colStorage=" & _
                                    OBJstate.usrState.objCollectionUtilities.collection2String _
                                    (colCollection, booReadable:=True)
                End If
#End If
            Case "LOOPEVENT"
                strString1 = objOperands(0).ToString
                strString2 = objOperands(1).ToString
                intInteger1 = CInt(objOperands(2))
                intInteger2 = CInt(objOperands(3))
                intInteger3 = CInt(objOperands(4))
                strString3 = objOperands(5).ToString
                RaiseEvent loopEvent(Me, _
                                     strString1, _
                                     strString2, _
                                     intInteger1, _
                                     intInteger2, _
                                     intInteger3, _
                                     strString3)
                If eventLoggingGet_() Then
                    strOperands = "strActivity=" & _OBJutilities.enquote(strString1) & vbNewLine & _
                                    "strEntity=" & _OBJutilities.enquote(strString2) & vbNewLine & _
                                    "intNumber=" & intInteger1 & vbNewLine & _
                                    "intCount=" & intInteger2 & vbNewLine & _
                                    "intLevel=" & intInteger3 & vbNewLine & _
                                    "strComment=" & _OBJutilities.enquote(strString3)
                End If
            Case "MSGEVENT"
                strString1 = objOperands(0).ToString
                RaiseEvent msgEvent(Me, strString1)
                If eventLoggingGet_() Then
                    strOperands = "strMessage=" & strString1
                End If
            Case "PARSEEVENT"
                strString1 = objOperands(0).ToString
                booBoolean1 = CBool(objOperands(1))
                intInteger1 = CInt(objOperands(2))
                intInteger2 = CInt(objOperands(3))
                intInteger3 = CInt(objOperands(4))
                intInteger4 = CInt(objOperands(5))
                intInteger5 = CInt(objOperands(6))
                intInteger6 = CInt(objOperands(7))
                strString2 = objOperands(8).ToString
                intInteger7 = CInt(objOperands(9))
                RaiseEvent parseEvent(Me, _
                                      strString1, _
                                      booBoolean1, _
                                      intInteger1, _
                                      intInteger2, _
                                      intInteger3, _
                                      intInteger4, _
                                      intInteger5, _
                                      intInteger6, _
                                      strString2, _
                                      intInteger7)
                If eventLoggingGet_() Then
                    strOperands = "strGrammarCategory=" & _OBJutilities.enquote(strString1) & vbNewLine & _
                                    "booTerminal=" & booBoolean1 & vbNewLine & _
                                    "intSrcStartInteger=" & intInteger1 & vbNewLine & _
                                    "intSrcLength=" & intInteger2 & vbNewLine & _
                                    "intTokStartIndex=" & intInteger3 & vbNewLine & _
                                    "intTokLength=" & intInteger4 & vbNewLine & _
                                    "intObjStartIndex=" & intInteger5 & vbNewLine & _
                                    "intObjLength=" & intInteger6 & vbNewLine & _
                                    "strComment=" & _OBJutilities.enquote(strString2) & vbNewLine & _
                                    "intLevel=" & intInteger7
                End If
            Case "PARSEFAILEVENT"
                RaiseEvent parseFailEvent(Me, CStr(objOperands(0)))
            Case "PARSESTARTEVENT"
                RaiseEvent parseStartEvent(Me, CStr(objOperands(0)))
            Case "SCANEVENT"
                Dim objToken As qbToken.qbToken = CType(objOperands(0), qbToken.qbToken)
                RaiseEvent scanEvent(Me, objToken)
                If eventLoggingGet_() Then
                    strOperands = "objToken=" & _OBJutilities.enquote(objToken.ToString)
                End If
            Case "THREADSTATUSEVENT"
                RaiseEvent threadStatusChangeEvent(Me)
            Case "USERERROREVENT"
                strString1 = objOperands(0).ToString
                RaiseEvent userErrorEvent(Me, strString1, strString2)
                If eventLoggingGet_() Then
                    strOperands = "strDescription=" & _OBJutilities.enquote(strString1) & vbNewLine & _
                                  "strHelp=" & strString2
                End If
            Case Else
                errorHandler_("Unknown event " & strName, _
                              "raiseEvent_", _
                              "Event not raised", _
                              Nothing)
                Dim intIndex1 As Integer
                For intIndex1 = 0 To UBound(objOperands)
                    _OBJutilities.append(strOperands, _
                                         vbNewLine, _
                                         "Operand" & CStr(intIndex1 + 1) & "=" & _
                                         _OBJutilities.object2String(objOperands(intIndex1)))
                Next intIndex1
        End Select
        If eventLoggingGet_() Then
            extendEventLog_(strName, strOperands)
        End If
    End Sub

    ' ----------------------------------------------------------------------
    ' Reconstruct Code
    '
    '
    Private Function rebuildCode_(ByVal objScanned As qbScanner.qbScanner, _
                                  ByVal strSourceCode As String, _
                                  ByVal intStartIndex As Integer, _
                                  ByVal intEndIndex As Integer) As String
        Dim intIndex1 As Integer
        Dim intIndexPrevious As Integer = 1
        Dim objStringBuilder As New Text.StringBuilder("")
        For intIndex1 = intStartIndex To intEndIndex
            With objScanned.QBToken(intIndex1)
                _OBJutilities.append(objStringBuilder, _
                                    _OBJutilities.copies(" ", .StartIndex - intIndexPrevious), _
                                    Mid(strSourceCode, .StartIndex, .Length))
                intIndexPrevious = .StartIndex + .Length
            End With
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Resets the quickBasicEngine
    '
    '
    Private Function reset_() As Boolean
        With OBJstate.usrState
            If Not (.objScanner Is Nothing) AndAlso Not clearScanner_(.objScanner) Then
                Return (False)
            End If
            If Not mkScanner_() Then Return (False)
            If Not resetCompiler_() Then Return (False)
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Resets a collection by clearing its old version (if it exists) and 
    ' creating a new version
    '
    '
    Private Function resetCollection_(ByRef colCollection As Collection, _
                                      ByVal strName As String) As Boolean
        If Not (colCollection Is Nothing) _
           AndAlso _
           Not OBJcollectionUtilities.collectionClear(colCollection) _
           OrElse _
           Not mkCollection_(colCollection, strName) Then
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Resets the compiler's data structures
    '
    ' 
    Private Function resetCompiler_() As Boolean
        With OBJstate.usrState
            If Not resetCollection_(.colConstantExpressionIndex, _
                                    "ConstantExpressionIndex") Then Return (False)
            If Not (.colPolish Is Nothing) Then
                If Not clearPolish_(.colPolish) Then Return (False)
                .colPolish = Nothing
            End If
            If Not mkCollection_(.colPolish, "RPN") Then Return (False)
            If Not resetCollection_(.colPolish, "RPN") Then Return (False)
            If Not resetCollection_(.colReadData, "Read Data") Then Return (False)
            If Not resetCollection_(.colSubFunctionIndex, _
                                    "subroutine and function index") Then Return (False)
            If Not (.colVariables Is Nothing) Then
                With .colVariables
                    Dim intIndex1 As Integer
                    For intIndex1 = 1 To .Count
                        If Not disposeVariable_(CType(.Item(intIndex1), _
                                                      qbVariable.qbVariable)) Then
                            Return (False)
                        End If
                    Next intIndex1
                End With
            End If
            If Not resetCollection_(.colVariables, "variables") Then Return (False)
            .booCompiled = False
            .booAssembled = False
            .enuSourceCodeType = ENUsourceCodeType.unknown
            .objImmediateResult = Nothing
            .booExplicit = False
            .intLabelSeq = 0
            Try
                ReDim .usrSubFunction(0)
                ReDim .usrConstantExpression(0)
            Catch objException As Exception
                errorHandler_("Not able to initialize or reallocate arrays: " & _
                              Err.Number & " " & Err.Description, _
                              "resetCompiler_", _
                              "Marking object unusable and returning False", _
                              objException)
                OBJstate.usrState.booUsable = False : Return (False)
            End Try
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Retrieve constant cache
    '
    '
    ' This method returns a string, a number or Nothing when the constant expression is
    ' not available.
    '
    '
    Private Function retrieveConstantCache_(ByVal strConstantExpression As String) _
            As qbVariable.qbVariable
        With OBJstate.usrState
            Try
                Dim colEntry As Collection = _
                CType(.colConstantExpressionIndex.Item(object2Key_(strConstantExpression)), _
                      Collection)
                Return (.usrConstantExpression(CInt(colEntry.Item(2))).objValue)
            Catch : End Try
            Return (Nothing)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Run code
    '
    '
    Private Function run_(ByVal strSourceCodeType As String) As Boolean
        Dim enuType As ENUsourceCodeType = string2SourceCodeType_(strSourceCodeType)
        If enuType = ENUsourceCodeType.invalid Then Return (False)
        Try
            executeCode_(OBJstate.usrState.objScanner.SourceCode, enuType)
        Catch objException As Exception
            errorHandler_("Error in running code: " & Err.Number & " " & Err.Description, _
                          "run_", _
                          "Returning False", _
                          objException)
            Return (False)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Change the source code: note that this resets the compiler
    '
    '
    Private Function sourceCodeSet_(ByVal strNew As String) As Boolean
        If Not reset_() Then Return (False)
        With OBJstate.usrState
            If (.objScanner Is Nothing) Then
                Try
                    .objScanner = New qbScanner.qbScanner
                Catch objException As Exception
                    errorHandler_("Cannot create scanner: " & _
                                  Err.Number & " " & Err.Description, _
                                  "", _
                                  "Marking object not usable: returning False", _
                                  objException)
                    .booUsable = False
                    Return (False)
                End Try
            End If
            .objScanner.SourceCode = strNew
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Convert source token index to surrounding statement
    '
    '
    ' When the optional overloaded parameter intEndIndex is not present
    ' this function returns the surrounding statement while assuming that
    ' the scan is complete, and its objScanned parameter contains all and
    ' only the source tokens.
    '
    ' The optional overloaded parameter intEndIndex is used when the 
    ' objScanned data structure is in the state it has during the scan,
    ' such that the current entry is the last entry: here, we return
    ' tokens back to the preceding newline or start of file, and 
    ' unscanned string characters up to the next newline in the source code.
    '
    '
    Private Overloads Function sourceIndex2Code_(ByVal intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByVal strSource As String) As String
        Return (sourceIndex2Code_(intIndex, objScanned, strSource, 0))
    End Function
    Private Overloads Function sourceIndex2Code_(ByVal intIndex As Integer, _
                                                 ByVal objScanned As qbScanner.qbScanner, _
                                                 ByVal strSource As String, _
                                                 ByVal intEndIndex As Integer) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intLength As Integer
        For intIndex1 = intIndex - 1 To 1 Step -1
            If objScanned.QBToken(intIndex1).TokenType _
               = _
               qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline Then Exit For
        Next intIndex1
        If intIndex1 < objScanned.TokenCount Then
            intIndex1 = objScanned.QBToken(intIndex1 + 1).StartIndex
        ElseIf intIndex1 = 0 Then
            intIndex1 = 1
        Else
            intIndex1 = Len(strSource) + 1
        End If
        Dim intEndIndexWork As Integer = CInt(IIf(intEndIndex = 0, objScanned.TokenCount, intEndIndex))
        For intIndex2 = intIndex To intEndIndexWork
            If objScanned.QBToken(intIndex2).TokenType _
               = _
               qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline Then
                intLength = objScanned.QBToken(intIndex2).StartIndex - intIndex1
                Exit For
            End If
        Next intIndex2
        If intIndex2 > intEndIndexWork Then
            If intEndIndex = 0 Then
                intLength = Len(strSource) + 1 - intIndex1
            Else
                intLength = InStr(objScanned.QBToken(intIndex).StartIndex, _
                                  strSource & vbNewLine, _
                                  vbNewLine) _
                            - _
                            intIndex1
            End If
        End If
        Return (Trim(Replace(Mid(strSource, intIndex1, intLength), vbNewLine, "\n")))
    End Function

    ' ----------------------------------------------------------------------
    ' Convert string to the source code type enumerator
    '
    '
    Private Function string2SourceCodeType_(ByVal strType As String) _
            As ENUsourceCodeType
        Select Case _OBJutilities.findAbbrev(strType, SOURCECODETYPELIST)
            Case 1 : Return (ENUsourceCodeType.unknown)
            Case 2 : Return (ENUsourceCodeType.immediateCommand)
            Case 3 : Return (ENUsourceCodeType.program)
            Case Else : Return (ENUsourceCodeType.invalid)
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Tester
    '
    '
    ' Note that this method expects TEST_CODE to be in this format:
    '
    '
    '      *  Each individual test case should be separated by a double
    '         newline
    '
    '      *  Two or three lines should be supplied for each case
    '
    '         + Line 1 should be the word "program" (if the test case
    '           evaluates a program) or the word "expression", followed
    '           by the program or expression.
    '
    '           Note that when a multi-line program is specified, line
    '           breaks should appear as the XML sequence &#00013&#00010.
    '
    '         + If line 1 is a program, line 2 must supply its input
    '           queue. This is the sequence of responses to Input commands
    '           separated by commas.
    '
    '         + Line 2 or 3 must be the expected results. If newlines should
    '           appear in the expected results they should be specified in
    '           XML notation as &#00013&#00010.
    '
    '
    ' Note that this method uses the Tag of the test engine to communicate the
    ' input queue to its input event handler, and to get print results, when
    ' it is testing programs.
    '
    ' This method Tags the test engine with a two item collection. Item(1) is
    ' a subcollection containing all members of the input queue for the test.
    ' Item(2) is set to a subcollection containing each distinct Printed ouput.
    '
    '
    Private Function test_(ByRef strReport As String, _
                           ByVal booEventLog As Boolean) As Boolean
        Dim strSplit() As String
        Try
            strSplit = Split(TEST_CODE, vbNewLine & vbNewLine)
        Catch objException As Exception
            errorHandler_("Can't perform a split: " & Err.Number & " " & Err.Description, _
                          "test", _
                          "Marking object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False : Return (False)
        End Try
        Try
            OBJtestQBE = New quickBasicEngine.qbQuickBasicEngine
        Catch objException As Exception
            errorHandler_("Can't create test engine: " & Err.Number & " " & Err.Description, _
                          "test", _
                          "Marking object unusable and returning False", _
                          objException)
            OBJstate.usrState.booUsable = False : Return (False)
        End Try
        AddHandler OBJtestQBE.interpretInputEvent, AddressOf test__inputEventHandler_
        AddHandler OBJtestQBE.interpretPrintEvent, AddressOf test__printEventHandler_
        With OBJtestQBE
            .EventLogging = booEventLog
            .ConstantFolding = OBJstate.usrState.booConstantFolding
            .DegenerateOpRemoval = OBJstate.usrState.booDegenerateOpRemoval
            .AssemblyRemovesCode = OBJstate.usrState.booAssemblyRemovesCode
            Dim booExpression As Boolean
            Dim booOK As Boolean = True
            Dim bytIndex1 As Byte
            Dim intCount As Integer
            Dim intIndex1 As Integer
            Dim strComment As String
            Dim strEventLog As String
            Dim strResult As String
            Dim strSplit2() As String
            Dim strTest As String
            Dim strTestDisplay As String
            Dim strTestType As String
            strReport = ""
            Dim datStart As Date = Now
            For intIndex1 = 0 To UBound(strSplit)
                If loopEventInterface_("Testing the Quick Basic engine", _
                                       "Test case", _
                                       intIndex1 + 1, _
                                       UBound(strSplit) + 1, _
                                       0, _
                                       "") Then Exit For
                intCount += 1
                Try
                    .reset()
                    strSplit2 = Split(strSplit(intIndex1), vbNewLine)
                    strTestType = _OBJutilities.word(strSplit2(0), 1)
                    strTest = _OBJutilities.phrase(strSplit2(0), 2)
                    strEventLog = ""
                    Select Case UCase(strTestType)
                        Case "EXPRESSION"
                            strResult = ""
                            Try
                                If booEventLog Then
                                    strResult = CStr(.evaluate(strTest, strEventLog).value)
                                Else
                                    strResult = CStr(.evaluate(strTest).value)
                                End If
                            Catch : End Try
                            bytIndex1 = 1
                            booExpression = True
                        Case "PROGRAM"
                            .Tag = test__mkTag_(strSplit2(1))
                            .SourceCode = Replace(strTest, "&#00013&#00010", vbNewLine)
                            .run()
                            bytIndex1 = 2
                            booExpression = False
                            strResult = test__outputQueue2String_(test__getCollection_(OBJtestQBE, 2))
                        Case Else
                            errorHandler_("Invalid test type in test script", _
                                          "test_", _
                                          "Returning False", _
                                          Nothing)
                            Return (False)
                    End Select
                    strResult = _OBJutilities.string2Display(strResult)
                    strComment = "expected result: " & _
                                 _OBJutilities.enquote(strSplit2(bytIndex1)) & ": " & _
                                 "actual result: " & _
                                 _OBJutilities.enquote(strResult)
                    If booExpression AndAlso booEventLog Then
                        strComment &= vbNewLine & vbNewLine & _
                                      _OBJutilities.string2Box(strEventLog, _
                                                               "Expression evaluation events")
                    End If
                    booOK = (strResult = strSplit2(bytIndex1))
                Catch
                    booOK = False
                    strComment = "error: " & Err.Number & " " & Err.Description
                End Try
                If booExpression Then
                    strTestDisplay = "Testing the expression " & _OBJutilities.enquote(strTest)
                Else
                    strTestDisplay = _OBJutilities.string2Box(.SourceCode, "Test Program")
                End If
                If booEventLog Then strEventLog = _OBJutilities.string2Box(.eventLogFormat, "Event Log")
                _OBJutilities.append(strReport, _
                                     vbNewLine & vbNewLine & vbNewLine, _
                                     strTestDisplay & _
                                     vbNewLine & vbNewLine & _
                                     strComment & _
                                     vbNewLine & vbNewLine & _
                                     strEventLog)
                If booEventLog Then
                    OBJstate.usrState.objCollectionUtilities.collectionClear(.EventLog)
                End If
                If Not booOK Then Exit For
            Next intIndex1
            Dim strPopCheck As String
#If QUICKBASICENGINE_POPCHECK Then
            strPopCheck = "Yes"
#Else
            strPopCheck = "No"
#End If
            strReport = "Quick Basic Engine Test" & "   " & Now & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        intCount & " test cases took " & _
                        DateDiff(DateInterval.Second, datStart, Now) & " " & _
                        "second(s)" & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        "Constant evaluation during parsing: " & _
                        CStr(IIf(.ConstantFolding, "Yes", "No")) & vbNewLine & _
                        "Degenerate operations removed: " & _
                        CStr(IIf(.ConstantFolding, "Yes", "No")) & vbNewLine & _
                        "Assembler removes labels and constants: " & _
                        CStr(IIf(.AssemblyRemovesCode, "Yes", "No")) & vbNewLine & _
                        "Stack is checked during interpretation: " & _
                        strPopCheck & vbNewLine & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        "The test " & CStr(IIf(booOK, "succeeded", "failed")) & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        strReport
            If Not .dispose(True) Then booOK = False
            Return (booOK)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Handle an Input event on behalf of the test method
    '
    '
    Private Sub test__inputEventHandler_(ByVal objQBE As qbQuickBasicEngine, _
                                         ByRef strChars As String)
        Dim colHandle As Collection = test__getCollection_(objQBE, 1)
        strChars = CStr(colHandle.Item(1))
        Try
            colHandle.Remove(1)
        Catch objException As Exception
            errorHandler_("Cannot remove element from the input queue", _
                          "test__inputEventHandler_", _
                          "Marking object unusable", _
                          objException)
            OBJstate.usrState.booUsable = False
        End Try
    End Sub

    ' ----------------------------------------------------------------------
    ' Make the tag for the test object as a collection
    '
    '
    Private Function test__mkTag_(ByVal strInputQueue As String) As Collection
        Try
            Dim colHandle As Collection = New Collection
            With colHandle
                .Add(New Collection)
                Dim strSplit() As String = Split(strInputQueue, ",")
                Dim intIndex1 As Integer
                With CType(.Item(1), Collection)
                    For intIndex1 = 0 To UBound(strSplit)
                        .Add(strSplit(intIndex1))
                    Next intIndex1
                End With
                .Add(New Collection)
            End With
            Return colHandle
        Catch objException As Exception
            errorHandler_("Cannot make Tag collection", _
                          "test__mkTag_", _
                          "Marking object unusable", _
                          objException)
            OBJstate.usrState.booUsable = False
            Return (Nothing)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Convert the output Print queue to a string on behalf of the test
    ' method
    '
    '
    Private Function test__outputQueue2String_(ByVal colQueue As Collection) As String
        With colQueue
            Dim intIndex1 As Integer
            Dim strOutstring As String
            For intIndex1 = 1 To .Count
                strOutstring &= CStr(.Item(intIndex1))
            Next intIndex1
            Return strOutstring
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Handle an Input event on behalf of the test method
    '
    '
    Private Sub test__printEventHandler_(ByVal objQBE As qbQuickBasicEngine, _
                                         ByVal strOutput As String)
        Dim colHandle As Collection = test__getCollection_(objQBE, 2)
        Try
            colHandle.Add(strOutput)
        Catch objException As Exception
            errorHandler_("Cannot update output print queue", _
                          "test__printEventHandler_", _
                          "Marking object unusable", _
                          objException)
            OBJstate.usrState.booUsable = False
        End Try
    End Sub

    ' ----------------------------------------------------------------------
    ' Get collection from tag
    '
    '
    Private Function test__getCollection_(ByVal objQBE As qbQuickBasicEngine, _
                                          ByVal bytIndex As Byte) _
            As Collection
        Return CType(CType(objQBE.Tag, Collection).Item(bytIndex), Collection)
    End Function

    ' ----------------------------------------------------------------------
    ' Trace code to structure
    '
    '
    Private Function traceCode2Struct_(ByVal intCode As Integer) As TYPtracing
        Dim usrTracing As TYPtracing
        With usrTracing
            Try
                Dim strBits As String = _OBJutilities.align(Mid(_OBJutilities.long2BaseN(intCode, "01"), 1, 16), 16, _OBJutilities.ENUalign.alignRight, "0")
                .booTextTrace = CBool(Mid(strBits, 16, 1))
                .booHeadsupTrace = CBool(Mid(strBits, 15, 1))
                .intRate = CInt(_OBJutilities.baseN2Long(Mid(strBits, 7, 8), "01"))
                .booString2Box = CBool(Mid(strBits, 6, 1))
                .booIncludeSource = CBool(Mid(strBits, 5, 1))
                .booIncludeMemory = CBool(Mid(strBits, 4, 1))
                .booIncludeStack = CBool(Mid(strBits, 3, 1))
                .booTraceLines = CBool(Mid(strBits, 2, 1))
                .booIncludeObject = CBool(Mid(strBits, 1, 1))
            Catch objException As Exception
                errorHandler_("Invalid trace type in trace code occured due to a programming error: trace cleared", _
                              "", _
                              "", _
                              objException)
                .booTextTrace = False
                .booHeadsupTrace = False
            End Try
            Return (usrTracing)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Trace type structure to numeric code
    '
    '
    ' The code returned uses the bits as follows (from the least significant
    ' position):
    '
    '
    '      *  Bit 0: text trace
    '      *  Bit 1: heads-up trace
    '      *  Bit 2-9: 8 bit rate (usually 1 for every line or every op)
    '      *  Bit 10: box trace elements
    '      *  Bit 11: include source lines
    '      *  Bit 12: include memory
    '      *  Bit 13: include stack
    '      *  Bit 15: trace in lines rather than opcodes
    '      *  Bit 16: include object code
    '
    '
    Private Function traceStruct2Code_(ByVal usrTracing As TYPtracing, _
                                       ByRef intIndex As Integer, _
                                       ByVal objScanned As qbScanner.qbScanner, _
                                       ByVal strSourceCode As String) As Integer
        With usrTracing
            Dim strRateByte As String = _OBJutilities.long2BaseN(CLng(.intRate), "01")
            If Len(strRateByte) > 8 Then
                compiler__errorHandler_("Trace rate " & .intRate & " is too large: limit is 255", _
                                        intIndex, _
                                        objScanned, _
                                        strSourceCode)
            End If
            Return CInt(_OBJutilities.baseN2Long(CStr(IIf(.booIncludeObject, "1", "0")) & _
                                   CStr(IIf(.booTraceLines, "1", "0")) & _
                                   CStr(IIf(.booIncludeStack, "1", "0")) & _
                                   CStr(IIf(.booIncludeMemory, "1", "0")) & _
                                   CStr(IIf(.booIncludeSource, "1", "0")) & _
                                   CStr(IIf(.booString2Box, "1", "0")) & _
                                   _OBJutilities.align(_OBJutilities.long2BaseN(CLng(.intRate), "01"), 8, _OBJutilities.ENUalign.alignRight, "0") & _
                                   CStr(IIf(.booHeadsupTrace, "1", "0")) & _
                                   CStr(IIf(.booTextTrace, "1", "0")), _
                                   "01"))
        End With
    End Function

    ' ------------------------------------------------------------------------
    ' Update cache of constant expressions and values: enforce maximum size
    '
    '
    Private Function updateConstantCache_(ByVal strConstantExpression As String, _
                                          ByVal objValue As qbVariable.qbVariable) As Boolean
        With OBJstate.usrState
            Dim objCacheValue As qbVariable.qbVariable = retrieveConstantCache_(strConstantExpression)
            If Not (objCacheValue Is Nothing) Then
                If Not objValue.compareTo(objCacheValue) Then
                    errorHandler_("Internal compiler error: " & _
                                    "different values developed for constant expression " & _
                                    _OBJutilities.enquote(strConstantExpression) & ": " & _
                                    "value in cache is " & _
                                    objCacheValue.ToString & ": " & _
                                    "value calculated is " & _
                                    objValue.ToString, _
                                    "updateConstantCache_", _
                                    "Marking object unusable and returning False", _
                                    Nothing)
                    .booUsable = False
                    Return False
                End If
                Return True
            End If
            Try
                Dim colEntry As Collection = New Collection
                colEntry.Add(strConstantExpression)
                colEntry.Add(UBound(.usrConstantExpression) + 1)
                .colConstantExpressionIndex.Add(colEntry, _
                                                object2Key_(strConstantExpression))
            Catch
            End Try
            Dim intUBound As Integer = UBound(.usrConstantExpression)
            Dim intExcessEntries As Integer = intUBound - CONSTANTEXPRESSION_MAXCACHE + 1
            Dim intIndex1 As Integer
            If intExcessEntries <= 0 Then
                Try
                    ReDim Preserve .usrConstantExpression(intUBound + 1)
                Catch objException As Exception
                    errorHandler_("Internal compiler error: " & _
                                  "cannot expand the constant expression table: " & _
                                  Err.Number & " " & Err.Description, _
                                  "", _
                                  "", _
                                  objException)
                    Return (False)
                End Try
                intIndex1 = UBound(.usrConstantExpression)
                If .intOldestConstantIndex = 0 Then .intOldestConstantIndex = 1
            Else
                intIndex1 = .intOldestConstantIndex
                .intOldestConstantIndex = (.intOldestConstantIndex + 1) _
                                          Mod _
                                          CONSTANTEXPRESSION_MAXCACHE
            End If
            With .usrConstantExpression(intIndex1)
                .strConstantExpression = strConstantExpression : .objValue = objValue
            End With
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Obtain the value of the variable
    '
    '
    Private Function variableGet_(ByVal objIndex As Object) As qbVariable.qbVariable
        Dim intIndex1 As Integer = derefVariableIndex_(objIndex)
        If intIndex1 = 0 Then Return Nothing
        Return CType(OBJstate.usrState.colVariables(intIndex1), _
                     qbVariable.qbVariable)
    End Function

    ' -----------------------------------------------------------------
    ' Modify the value of the variable
    '
    '
    Private Function variableSet_(ByVal objIndex As Object, _
                                  ByVal objValue As qbVariable.qbVariable) _
            As Boolean
        With OBJstate.usrState.colVariables
            Dim intIndex1 As Integer = derefVariableIndex_(objIndex)
            If intIndex1 = 0 Then Return False
            Try
                .Remove(intIndex1)
                .Add(objValue, objValue.VariableName, intIndex1)
            Catch
                errorHandler_("Cannot replace variable: " & _
                              Err.Number & " " & Err.Description, _
                              "variableSet_", _
                              "Object has been marked unusable", _
                              Nothing)
                OBJstate.usrState.booUsable = False
                Return False
            End Try
            Return True
        End With
    End Function

    ' ***** Event handlers ********************************************

    Private Sub scanErrorEventHandler(ByVal strMsg As String, _
                                      ByVal intIndex As Integer, _
                                      ByVal intLineNumber As Integer, _
                                      ByVal strHelp As String) _
            Handles OBJscanner.scanErrorEvent
        raiseEvent_("compileErrorEvent", _
                    strMsg, _
                    intIndex, _
                    1, _
                    intLineNumber, _
                    strHelp, _
                    OBJstate.usrState.objScanner.SourceCode)
    End Sub

    Private Sub scanEventHandler(ByVal objQBtoken As qbToken.qbToken, _
                                 ByVal intCharacterIndex As Integer, _
                                 ByVal intLength As Integer, _
                                 ByVal intTokenCount As Integer) _
            Handles OBJscanner.scanEvent
        raiseEvent_("scanEvent", objQBtoken)
        loopEventInterface_("Scanning", _
                            "character", _
                            intCharacterIndex, _
                            intLength, _
                            0, _
                            "Number of tokens: " & intTokenCount)
    End Sub

    Private Sub threadStatusEventHandler() Handles OBJthreadStatus.statusChangeEvent
        raiseEvent_("threadStatusEvent")
    End Sub

    Private Class threadStatus

        ' *********************************************************************
        ' *                                                                   *
        ' * threadStatus                                                      *
        ' *                                                                   *
        ' *                                                                   *
        ' * This class maintains the threading status of the Quick Basic      *
        ' * Engine instance, and it exposes the following methods, properties,*
        ' * and events.                                                       *
        ' *                                                                   *
        ' *                                                                   *
        ' *      *  available: this method returns True if the threadStatus is*
        ' *         Running or Ready, False otherwise.                        *
        ' *                                                                   *
        ' *      *  endThread: this method is called at the end of a          *
        ' *         thread's execution of a Public procedure. It decrements   *
        ' *         the thread count.                                         *
        ' *                                                                   *
        ' *      *  getThreadStatus: this method returns the thread status as *
        ' *         an enumerator of type ENUthreadStatus or as a string.     *
        ' *                                                                   *
        ' *         When called with no operands, getThreadStatus returns one *
        ' *         of the ENUthreadStatus values Ready, Running, Stopping,   *
        ' *         Stopped or Initializing.                                  *
        ' *                                                                   *
        ' *         When called with the Boolean operator True,               *
        ' *         getThreadStatus returns one of the strings Ready, Running,*
        ' *         Stopping, Stopped or Initializing.                        *
        ' *                                                                   *
        ' *      *  mkReady: when the threadStatus object is created it is in *
        ' *         the Initializing status: this method places the new       *
        ' *         status in the Ready state.                                *
        ' *                                                                   *
        ' *      *  runningThreads: this method returns the count of running  *
        ' *         threads. Note that this method will always return at least*
        ' *         one when it is called in a Quick Basic Engine procedure   *
        ' *         that uses our startThread to indicate that a thread is    *
        ' *         starting up.                                              *
        ' *                                                                   *
        ' *      *  setThreadStatus(e): this method sets the thread status to *
        ' *         the ENUthreadStatus passed in e.                          *
        ' *                                                                   *
        ' *      *  startThread: this method is called at the start of a      *
        ' *         thread's execution of a Public procedure. If a thread     *
        ' *         may be started (if the object is not Stopped or Stopping) *
        ' *         this method adds one to the thread count and returns      *
        ' *         True. Otherwise this method returns False.                *
        ' *                                                                   *
        ' *      *  statusChangeEvent: this event is raised when the thread   *
        ' *         status changes                                            *
        ' *                                                                   *
        ' *      *  stopThread: this method decrements the thread count.      *
        ' *                                                                   *
        ' *      *  stopThreads: this method initiates the stopping process.  *
        ' *         If no thread beyond the current thread is executing, this *
        ' *         method places the threadStatus (and, in effect, the Quick *
        ' *         Basic engine) into the Stopped state.                     *
        ' *                                                                   *
        ' *         If more than one thread is executing this method places   *
        ' *         the runStatus and the QBE into the Stopping state.        *
        ' *                                                                   *
        ' *      *  toString: this method returns the threadStatus as a string*
        ' *                                                                   *
        ' *********************************************************************

        ' ***** Thread status as an enumerator *****
        Public Enum ENUthreadStatus
            Ready                       ' No threads running except enquiry thread
            Stopped                     ' Stop has been executed and nothing is running 
            Stopping                    ' Stop has been executed and something is running 
            Running                     ' One or more threads running beyond enquiry
            Initializing                ' Startup state
        End Enum

        ' ***** Object state *****
        Public Structure TYPstate
            Dim intThreadCount As Integer   ' Number of threads
            Dim booStopping As Boolean      ' True: stopped when intThreadCount=0:
                                            ' stopping when intThreadCount>0
            Dim booInitializing As Boolean  ' Starting up
        End Structure
        Private USRstate As TYPstate

        ' ***** Events *****
        Public Event statusChangeEvent()

        ' ---------------------------------------------------------------------
        ' Object constructor
        '
        ' 
        Public Sub New()
            USRstate.booInitializing = True
        End Sub

        ' ---------------------------------------------------------------------
        ' Return availability
        '
        '
        Public Function available() As Boolean
            With Me
                Dim enuThreadStatusWork As ENUthreadStatus = Me.getThreadStatus
                Return (enuThreadStatusWork = ENUthreadStatus.Ready _
                        OrElse _
                        enuThreadStatusWork = ENUthreadStatus.Running)
            End With
        End Function

        ' ---------------------------------------------------------------------
        ' End one thread
        '
        '
        Public Function endThread() As Boolean
            USRstate.intThreadCount -= 1
            RaiseEvent statusChangeEvent()
            Return (True)
        End Function

        ' ---------------------------------------------------------------------
        ' Get thread status
        '
        ' 
        ' --- Returns the status as a string
        Public Function getThreadStatus(ByVal booFlag As Boolean) As String
            Return (Me.getThreadStatus.ToString)
        End Function
        ' --- Returns the status as an enum
        Public Function getThreadStatus() As ENUthreadStatus
            With USRstate
                If .booInitializing Then Return ENUthreadStatus.Initializing
                If .intThreadCount = 0 Then
                    Return (CType(IIf(.booStopping, _
                                      ENUthreadStatus.Stopped, _
                                      ENUthreadStatus.Ready), _
                                  ENUthreadStatus))
                Else
                    Return (CType(IIf(.booStopping, _
                                      ENUthreadStatus.Stopping, _
                                      ENUthreadStatus.Running), _
                                  ENUthreadStatus))
                End If
            End With
        End Function

        ' ---------------------------------------------------------------------
        ' Move to the ready status from the initialising status
        '
        ' 
        Public Function mkReady() As Boolean
            USRstate.booInitializing = False
        End Function

        ' ---------------------------------------------------------------------
        ' Get count of running threads
        '
        ' 
        Public Function runningThreads() As Integer
            Return (USRstate.intThreadCount)
        End Function

        ' ---------------------------------------------------------------------
        ' Change the thread status
        '
        '
        Public Function setThreadStatus(ByVal enuNew As ENUthreadStatus) _
               As Boolean
            With USRstate
                Select Case enuNew
                    Case ENUthreadStatus.Initializing
                        .booInitializing = True
                        .booStopping = False
                        .intThreadCount = 0
                    Case ENUthreadStatus.Ready
                        .booInitializing = False
                        .booStopping = False
                        .intThreadCount = 0
                    Case ENUthreadStatus.Stopping
                        .booInitializing = False
                        .booStopping = True
                    Case Else
                        utilities.utilities.errorHandler("Invalid new status " & _
                                                         enuNew.ToString, _
                                                         "threadStatus", _
                                                         "setThreadStatus", _
                                                         "No change made")
                        Return (False)
                End Select
                Return (True)
            End With
        End Function

        ' ---------------------------------------------------------------------
        ' "Start" a thread if possible
        '
        ' 
        Public Function startThread() As Boolean
            If Me.getThreadStatus = ENUthreadStatus.Ready _
                OrElse _
                Me.getThreadStatus = ENUthreadStatus.Running Then
                USRstate.intThreadCount += 1
                RaiseEvent statusChangeEvent()
                Return (True)
            End If
            Return (False)
        End Function

        ' ---------------------------------------------------------------------
        ' Stop thread 
        '
        ' 
        Public Function stopThread() As Boolean
            With USRstate
                .intThreadCount -= 1
                If .intThreadCount < 0 Then .intThreadCount = 0
                RaiseEvent statusChangeEvent()
                Return (True)
            End With
        End Function

        ' ---------------------------------------------------------------------
        ' Stop all threads
        '
        ' 
        Public Function stopThreads() As Boolean
            USRstate.booStopping = True
            RaiseEvent statusChangeEvent()
            Return (True)
        End Function

        ' ---------------------------------------------------------------------
        ' Convert to string
        '
        ' 
        Public Overrides Function toString() As String
            Return (Me.getThreadStatus.ToString & ": " & _
                    Me.runningThreads & " thread(s) are running")
        End Function

    End Class

#Region " Internal event handlers "

    Private Sub compileErrorEvent_handler_(ByVal objQBsender As qbQuickBasicEngine, _
                                           ByVal strMessage As String, _
                                           ByVal intIndex As Integer, _
                                           ByVal intContextLength As Integer, _
                                           ByVal intLinenumber As Integer, _
                                           ByVal strHelp As String, _
                                           ByVal strCode As String)
        RaiseEvent msgEvent(Me, _
                            "Note that the error " & _
                            _OBJutilities.enquote(_OBJutilities.ellipsis(strMessage, 32)) & " " & _
                            "occured but may be transitory " & _
                            "because the input code is being checked for an immediate " & _
                            "command")
    End Sub

#End Region

End Class

