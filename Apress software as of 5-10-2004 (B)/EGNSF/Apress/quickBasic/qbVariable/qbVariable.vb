Option Strict On

Imports System.Threading

Imports qbVariableType.qbVariableType

Imports qbTokenType.qbTokenType

Public Class qbVariable

    Implements IComparable
    Implements ICloneable
    Implements IDisposable

    ' ***********************************************************************
    ' *                                                                     *
    ' * quickBasicEngine Variable                                           *
    ' *                                                                     *
    ' *                                                                     *
    ' * This class represents the type, structure and value of a Quick Basic*
    ' * scalar or an n-dimensional Quick Basic array.                       *
    ' *                                                                     *
    ' * The rest of this block describes:                                   *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  The variable data model                                     *
    ' *      *  The qbVariable state                                        *
    ' *      *  Containment, isomorphism, and string identity of variables  *
    ' *      *  Properties, methods and events of this class                *
    ' *      *  fromString BNF syntax                                       * 
    ' *      *  Multithreading considerations                               * 
    ' *                                                                     *
    ' *                                                                     *
    ' * THE VARIABLE DATA MODEL ------------------------------------------- *
    ' *                                                                     *
    ' * As far as we're concerned, a Quick Basic variable can be a scalar   *
    ' * (of type Boolean, Byte, Integer, Long, Single, Double, or String),  *
    ' * a Variant that contains (has-a) nonvariant data item, an Array as   *
    ' * a collection of entries with a nonvariant type, a User Data Type,   *
    ' * also represented as a collection of entries, or an Unknown value    *
    ' * or Null (represented by Nothing in .Net.)                           *
    ' *                                                                     *
    ' * In all cases a variable has structure and data.                     *
    ' *                                                                     *
    ' * The structure of a scalar is just its type.                         *
    ' *                                                                     *
    ' * The structure of a Variant is the type (not including the value) of *
    ' * what it contains. The structure of an ordinary Variant is "concrete"*
    ' * because it cannot be specified unless the accompanying contained    *
    ' * non-variant type is also specified.                                 *
    ' *                                                                     *
    ' * The structure of an array is its number of dimensions, and, in Quick*
    ' * Basic, upper and lower bounds, which can be almost any positive or  *
    ' * negative numbers (the only semi-useful ability to start an array at *
    ' *  a lower bound, other than one, was dropped by .Net). This structure*
    ' * is sometimes referred to as a dope vector and in these comments I   *
    ' * will generalize the "dope concept" to use dope as a synonym for the *
    ' * structure of any variable.                                          *
    ' *                                                                     *
    ' * The structure of an array also includes the uniform (what we refer  *
    ' * to as orthogonal) type of all the data in the array. This type can  *
    ' * be an "abstract" Variant type which as a pure Variant type doesn't  *
    ' * specify (indeed, may not specify) the contained type, because in an *
    ' * array this will vary.                                               *
    ' *                                                                     *
    ' * The structure of a user defined type is the ordered collection of   *
    ' * variables. Note that in Quick Basic, this collection can't be       *
    ' * nested. It cannot contain, directly, definitions of further UDTs    *
    ' * (as seen, for example, in Cobol). But, it can contain UDTs defined  *
    ' * elsewhere.                                                          *
    ' *                                                                     *
    ' * The structure of an Unknown or Null variable is just its being-     *
    ' * Unknown or its being-Nothing.                                       *
    ' *                                                                     *
    ' * Data is the data associated with the variable. For a scalar, this is*
    ' * represented by the corresponding .Net type with 2 important         *
    ' * exceptions: Quick Basic Integers are represented by .Net Short      *
    ' * integers because .Net Integers are 32-bit while Quick Basic         *
    ' * Integers are 16-bit: Quick Basic Longs are represented by .Net      *
    ' * Integers because .Net Longs are 64-bit while Quick Basic Integers   *
    ' * are 32-bit.                                                         *
    ' *                                                                     *
    ' * Note that mapping of Quick Basic variables to .Net variables is     *
    ' * accurate UNLESS the variable is Single or Double. Single and double *
    ' * precision is not at this writing mapped accurately and single and   *
    ' * double values will have a wider range than the corresponding Quick  *
    ' * Basic values. This means that numerical results may differ when     *
    ' * running old Quick Basic code using this object and the              *
    ' * quickBasicEngine.                                                   *
    ' *                                                                     *
    ' * 1-dimensional array data is represented by a collection of .Net     *
    ' * objects. 2-dimensional arrays are represented by a collection that  *
    ' * contains 1 or more "orthogonal" member collections, "orthogonal"    *
    ' * in the sense that each subcollection has an identical number of     *
    ' * members. In general, n-dimensional arrays are represented by fully  *
    ' * orthogonal balanced trees of subcollections.                        *
    ' *                                                                     *
    ' * Variants actually contain an instance of this type (qbVariable) in  *
    ' * their state which provides the type and the data of the variant     *
    ' * value. This qbVariable, however, is prevented from being itself a   *
    ' * Variant: but note that we do allow it to be an array.               *
    ' *                                                                     *
    ' * User data types contain a collection of qbVariables representing    *
    ' * their components.                                                   *
    ' *                                                                     *
    ' * THE QBVARIABLE STATE ---------------------------------------------- *
    ' *                                                                     *
    ' * The state of this object consists of the following:                 *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  booUsable: to indicate the object instance's usability      *
    ' *                                                                     *
    ' *      *  strName: the object instance name                           *
    ' *                                                                     *
    ' *      *  objDope: the variable structure represented as a            *
    ' *         qbVariableType.                                             *
    ' *                                                                     *
    ' *      *  objValue: the .Net representation of the variable value.    *
    ' *         This will be one of the following:                          *
    ' *                                                                     *
    ' *              + qbVariable is scalar type: objValue will be a .Net   *
    ' *                value of type Boolean, Short, Integer, Single,       *
    ' *                Double or String                                     *
    ' *                                                                     *
    ' *              + qbVariable is Variant type: objValue will be a       *
    ' *                qbVariable which won't have Variant type             *
    ' *                                                                     *
    ' *              + qbVariable is Array type: objValue will be a         *
    ' *                Collection structured orthogonally as described in   *
    ' *                the previous section                                 *
    ' *                                                                     *
    ' *              + qbVariable is UDT type: objValue will be a           *
    ' *                Collection of qbVariables, one for each member.      *
    ' *                                                                     *
    ' *              + qbVariable is Unknown or Null type: objValue will be *
    ' *                Nothing                                              *
    ' *                                                                     *
    ' *                                                                     *
    ' * CONTAINMENT, ISOMORPHISM AND STRING IDENTITY OF VARIABLES --------- *
    ' *                                                                     *
    ' * The containedVariable, isomorphicVariable and stringIdentical       *
    ' * methods support comparision of variables.                           *
    ' *                                                                     *
    ' * The variable a is "contained in" the variable b solely by virtue    *
    ' * of its underlying type; if all potential values of a can be         *
    ' * assigned to b then a is contained in b. For more on this, see the   *
    ' * documentation header of the qbVariableType class.                   *
    ' *                                                                     *
    ' * The variable a is isomorphic to b when the type of a is contained   *
    ' * in b, the type of b is contained in a, and the values of a and b    *
    ' * (after conversion to a string) are the same.                        *
    ' *                                                                     *
    ' * The conversion to a string is performed by calling toString, the    *
    ' * method described in the next section for serializing the variable   *
    ' * to type:value(s), and throwing away the material to the left of the *
    ' * colon as well as the colon. The stringIdentical method will test two*
    ' * qbVariable objects for this type of identity.                       *
    ' *                                                                     *
    ' * Note that in the case of scalar isomorphism, a and b will have      *
    ' * identical type and identical value. But if a and b are arrays, then *
    ' * they may differ in lowerBounds and upperBounds while retaining all  *
    ' * other common properties.                                            *
    ' *                                                                     *
    ' * UDTs at this writing are NEVER contained or isomorphic.             *
    ' *                                                                     *
    ' *                                                                     *
    ' * THE PROPERTIES, METHODS AND EVENTS OF THIS CLASS ------------------ *
    ' *                                                                     *
    ' * Properties are named starting with an upper case letter; methods    *
    ' * and events use camelCase and start with a lower case letter. Events *
    ' * are named ending in Event or EventShared.                           *
    ' *                                                                     *
    ' * For reliability and security, the object is at all times, upon the  *
    ' * completion of the object constructor, either Usable or not Usable.  *
    ' *                                                                     *
    ' * Upon completion of a successful New constructor, the object is      *
    ' * Usable until a serious error occurs, the mkUnusable method is run,  *
    ' * or the dispose method is executed.  Most procedures will then check *
    ' * usability, and refuse to proceed when the object is unusable.       *
    ' *                                                                     *
    ' * All methods, which do not otherwise return a value, return when     *
    ' * called as functions a Boolean success flag.                         *
    ' *                                                                     *
    ' * This class implements the ICompareable, ICloneable and IDisposable  *
    ' * interfaces: see the compareTo, Clone, dispose and disposeInspect    *
    ' * methods, below.                                                     *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  The Shared, read-only property About returns information    *
    ' *         identifying this class.                                     *
    ' *                                                                     *
    ' *      *  The Shared class2XML method returns the information about   *
    ' *         the class; note that it can be used in place of object2XML  *
    ' *         on an uninstantiated object to get some XML information     *
    ' *         out of the stateless code.                                  *
    ' *                                                                     *
    ' *      *  The Shared, read-only property ClassName returns the class  *
    ' *         name (qbVariable).                                          *
    ' *                                                                     *
    ' *      *  The clearVariable method clears the variable. If it is a    *
    ' *         scalar it is set to the default appropriate to its type,    *
    ' *         which is False for Booleans, 0 for numeric types, and a     *
    ' *         null string for strings. If the variable is a Variant, it   *
    ' *         is set to the default appropriate to its contained type.    *
    ' *                                                                     *
    ' *         If the variable is an Array, each entry is set to the       *
    ' *         appropriate default. If the variable is a UDT each member   *
    ' *         is cleared according to its type. Finally, if the variable  *
    ' *         is Unknown or Null no change is made.                       *
    ' *                                                                     *
    ' *         See also resetVariable.                                     *
    ' *                                                                     *
    ' *      *  The clone method creates a new copy of the object instance  *
    ' *         with identical values, returning this copy as its function  *
    ' *         value.                                                      *
    ' *                                                                     *
    ' *      *  The compareTo method returns True when the object matches   *
    ' *         another instance. If both instances are usable, and they    *
    ' *         have identical toString values, the compareTo method returns*
    ' *         True; otherwise, it returns False.                          *
    ' *                                                                     *
    ' *      *  The containedVariable(v) method returns True when the       *
    ' *         qbVariable object in v is contained in the instance as      *
    ' *         described in the preceding section.                         *
    ' *                                                                     *
    ' *         The overloaded syntax containedVariable(v,s) can be used to *
    ' *         derive the explanation as to why v is (or isn't) contained  *
    ' *         in the instance. s is a string passed by reference.         *
    ' *                                                                     *
    ' *         Note: for any two UDTs containedVariable will always return *
    ' *         False.                                                      *
    ' *                                                                     *
    ' *      *  The derefMemberName(n) method is valid only for variables   *
    ' *         that are UDTs. In this method, n should be the name of a UDT*
    ' *         member, and this method returns the qbVariable object,      *
    ' *         contained directly or indirectly in the overall instance,   *
    ' *         identified by n.                                            *
    ' *                                                                     *
    ' *         n may be a simple member name. If it selects a member that  *
    ' *         is a UDT, n may be simple in which case it returns the udt. *
    ' *         For a UDT n may also select sub-members when periods        *
    ' *         separate names: for example, when a udt contains UDT01, and *
    ' *         UDT01 contains intVal, then derefMemberName("UDT01.intVal") *
    ' *         returns the object corresponding to intVal.                 *   
    ' *                                                                     *
    ' *      *  The dispose method disposes the heap storage associated with*
    ' *         the object (if any) and marks the object not usable.  For   *
    ' *         best results, use this method (or disposeInspect) when you  *
    ' *         are finished with the object.                               *
    ' *                                                                     *
    ' *         See also disposeInspect.                                    *
    ' *                                                                     *
    ' *      *  The disposeInspect method disposes the heap storage         *
    ' *         associated with the object (if any) and marks the object not*
    ' *         usable. This method also will conduct a final inspection for*
    ' *         internal errors. For best results, use this method (or      *
    ' *         dispose) when you are are finished with the object.         *
    ' *                                                                     *
    ' *         See also dispose.                                           *
    ' *                                                                     *
    ' *      *  The read-write property Dope returns and can change         *
    ' *         info about the variable as an instance of the               *
    ' *         class qbVariableType.                                       *
    ' *                                                                     *
    ' *         Note: the default Dope is the Unknown qbVariableType        *
    ' *                                                                     *
    ' *         This property may be set to any qbVariableType. Changing    *
    ' *         this property usually clears the variable: if the           *
    ' *         variable is scalar or variant, it is set to its appropriate *
    ' *         default: if the variable is an array or UDT, each entry is  *
    ' *         set to its default.                                         *
    ' *                                                                     *
    ' *         However, when an array structure is changed to an isomorphic*
    ' *         structure (same dimensions, identical element types, and    *
    ' *         same size at each dimension), setting Dope won't clear the  *
    ' *         array. Otherwise the array is cleared.                      *
    ' *                                                                     *
    ' *      *  The empiricalDope method returns a reconstruction of the    *
    ' *         variable's type (including array bounds when the variable   *
    ' *         is an array) from its data exclusively. This reconstruction *
    ' *         is returned as a string acceptable to the fromString method *
    ' *         of qbVariableType.                                          *
    ' *                                                                     *
    ' *         If the variable is NOT an array, the empiricalDope will be  *
    ' *         identical to the variable's dope. If the variable IS an     *
    ' *         array, the empiricalDope will be isomorphic to the array    *
    ' *         dope of the variable; dimensions and bound sizes will be the*
    ' *         same, as well as the entry type; but lowerBounds of the     *
    ' *         empiricalDope will be zero.                                 *
    ' *                                                                     *
    ' *         Since a valid instance contains type information, this      *
    ' *         method is primarily a curiosa, for internal use and to      *
    ' *         clarify the concept of deriving a type from data only, which*
    ' *         we need when changing the data of an array without,         *
    ' *         unnecessarily, changing its structure.                      *
    ' *                                                                     *
    ' *      *  The fromString method assigns both the dope and values of   *
    ' *         the array. fromString(s) assigns the dope and values using a*
    ' *         string in the form type:value.                              * 
    ' *                                                                     *
    ' *         + type should be the variable type in the syntax supported  *
    ' *           by qbVariableType.fromString, and one of the following.   *
    ' *                                                                     *
    ' *                - For a scalar, type should be scalarType, where     *
    ' *                  scalarType is one of Boolean, Byte, Integer, Long, *
    ' *                  Single, Double or String.                          *
    ' *                                                                     *
    ' *                - For a Variant, type should be Variant,<scalarType>,*
    ' *                  Variant,(<arrayType>), or Variant,(<udt>). The     *
    ' *                  scalar type should be as described above; the      *
    ' *                  arrayType or udt should be in parentheses and as   *
    ' *                  described below.                                   *
    ' *                                                                     *
    ' *                - For an Array, type should be Array,type,bounds,    *
    ' *                  where type is one of the following:                *
    ' *                                                                     *
    ' *                  . The name of a scalar type                        *
    ' *                                                                     *
    ' *                  . The keyword Variant; note that the type of       *
    ' *                    the entries in variant arrays is specified       *
    ' *                    "abstractly" and with no associated scalar type  *
    ' *                                                                     *
    ' *                  . A parenthesized user data type definition        *
    ' *                                                                     *
    ' *                - For a user data type, the type should be UDT,      *
    ' *                  memberlist, where the member list consists of one  *
    ' *                  or more comma-separated and parenthesized member   *
    ' *                  definitions.                                       *
    ' *                                                                     *
    ' *                  Each definition has the form (name,type), where    *
    ' *                  name is the member name and type is its type. The  *
    ' *                  type must be scalar, abstract variant or array.    *
    ' *                                                                     *
    ' *                - For the Unknown type, type should be Unknown: for  *
    ' *                  Null is should be Null.                            *
    ' *                                                                     *
    ' *         + "values" should specify the variable value(s).            *
    ' *                                                                     *
    ' *                - If the variable is a scalar or a variant that does *
    ' *                  not contain an array, the value may be the         *
    ' *                  scalar's value (compatible with its type) as True, *
    ' *                  False, a number, or a string, quoted using Visual  *
    ' *                  Basic's conventions.                               *
    ' *                                                                     *
    ' *                  Alternatively the variable may be in "decorated"   *
    ' *                  form as type(value) where the value is True, False,*
    ' *                  a number or a string.                              *
    ' *                                                                     *
    ' *                  The variable may be represented as an asterisk.    *
    ' *                  This will assign the appropriate default value for *
    ' *                  the type.                                          *
    ' *                                                                     *
    ' *                - If the variable is an array, the value should be   *
    ' *                  the list of array values.                          *
    ' *                                                                     *
    ' *                  This is a comma-separated list of scalar values    *
    ' *                  (plain or "decorated") for a one dimensional array.*
    ' *                                                                     *
    ' *                  This is a comma-separated list of parenthesized    *
    ' *                  rows for a two dimensional array.                  *
    ' *                                                                     *
    ' *                  In general this is a comma-separated list of       *
    ' *                  array "slices" (arrays of one dimension lower)     *
    ' *                  for n dimensional arrays.                          *
    ' *                                                                     *
    ' *                  Each value in the array may optionally be followed *
    ' *                  by a repeat count in parentheses. The entry value  *
    ' *                  will be repeated until the end of the current      *
    ' *                  slice or the indicated number of entries. The      *
    ' *                  repeat count may be an asterisk to repeat to the   *
    ' *                  end of the current array slice.                    *
    ' *                                                                     *
    ' *                  Note that when the variable is not otherwise       *
    ' *                  known to be an array (when, for instance, the      *
    ' *                  variable type is omitted from the fromString       *
    ' *                  expression, the use of a repeat count will make the*
    ' *                  variable into an array.                            *
    ' *                                                                     *
    ' *                - For a UDT the value should be the comma separated  *
    ' *                  list of member values. Each member that is a scalar*
    ' *                  or the scalar value of a variant member should be  *
    ' *                  its value in string form or in the decorated form  *
    ' *                  type(value). Each member that is an array should be*
    ' *                  the array's value, represented "orthogonally" as   *
    ' *                  described above and enclosed in parentheses.       *  
    ' *                                                                     *
    ' *                  Each member that is a UDT should be the NESTED     *
    ' *                  UDT specification, in parentheses.                 *
    ' *                                                                     *
    ' *                - For Unknown and Null types "values" (and its       *
    ' *                  preceding column) should not be specified.         *
    ' *                                                                     *
    ' *         If the type is omitted and the value is specified, the type *
    ' *         will default to the narrowest Quick Basic type capable of   *
    ' *         containing the specified data. If the type is present and   *
    ' *         the value is omitted, the the variable will take on the     *
    ' *         default contents for the type.                              *
    ' *                                                                     *
    ' *         The syntax :value (no type, colon, value) may be used to    *
    ' *         change the value of the variable without altering the type. *
    ' *         The value must be compatible with the existing type, unless *
    ' *         the existing type is Unknown; in this case, the type will   *
    ' *         be changed, to the narrowest QB type capable of containing  *
    ' *         the value.                                                  *
    ' *                                                                     *
    ' *         Here are examples of valid fromString strings.              *
    ' *                                                                     *
    ' *         + Integer:4 specifies a 16-bit integer containing the value *
    ' *           four.                                                     * 
    ' *                                                                     *
    ' *         + Variant,Integer:4 specifies a variant that contains a     *
    ' *           16-bit integer containing the value four.                 * 
    ' *                                                                     *
    ' *         + Array,Integer,0,3:1,2,3,4 specifies an integer array      *
    ' *                                                                     *
    ' *         + Array,Integer,1,2,1,2:(1,2),(1,2) specifies a 2-dimension *
    ' *           integer matrix                                            *
    ' *                                                                     *
    ' *         + Array,Variant,1,2:Integer(1),Long(2) specifies a          *
    ' *           1-dimension variant array and it uses "decoration" to be  *
    ' *           specific about type.                                      *
    ' *                                                                     *
    ' *         + 32768 by  itself specifies a Long integer containing      *
    ' *           32768.                                                    *
    ' *                                                                     *
    ' *         + :32767 will assign 32767 to a pre-specified type. When    *
    ' *           set after the above example, :32767 will preserve the     *
    ' *           type of Long integer. When assigned to an uninitialized   *
    ' *           variable, :32767 creates a 16-bit integer.                *                                  
    ' *                                                                     *
    ' *         + Array,Byte,0,1 by  itself specifies a Byte array that     *
    ' *           contains the Byte default values of 0. The toString will  *
    ' *           be Array,Byte,0,1:*                                       *
    ' *                                                                     *
    ' *         + Array,Byte,0,1:*,1 specifies a Byte array that contains   *
    ' *           the Byte default value of 0 followed by 1. The toString   *
    ' *           will be Array,Byte,0,1:*,1                                *
    ' *                                                                     *
    ' *         + Array,Variant,0,1,1,2:(32767,"B"),(32768,1) specifies a   *
    ' *           Variant array. The toString will be                       *
    ' *           Array,Variant,0,1,1,2:(System.Int16(32767),               *
    ' *           System.String("B")),(System.Int32(32768),System.Byte(1)), *
    ' *           and note that values are decorated, because the array has *
    ' *           variant entries.                                          *
    ' *                                                                     *
    ' *         + UDT,(intMember01,Integer),(strMember02,Array,String,1,2), *
    ' *           (typMember03,(udt,(intMember01,Integer))):1,("A","B"),    *
    ' *           (udt,(intMember01,Integer):1) specifies a UDT containing  *
    ' *           an integer, a string array, and an inner UDT.             *
    ' *                                                                     *
    ' *         In the fromString string, either the type or the value of   *
    ' *         the variable can be omitted.                                *
    ' *                                                                     *
    ' *         If the type is omitted, it is determined "empirically" by   *
    ' *         examining the value string:                                 *
    ' *                                                                     *
    ' *         + If the value string is null or an asterisk then the type  *
    ' *           is Unknown                                                *
    ' *                                                                     *
    ' *         + If the value string is a number or quoted string then     *
    ' *           the type is the narrowest scalar Quick Basic type that    *
    ' *           can contain the value                                     *
    ' *                                                                     *
    ' *         + If the value string is in parentheses OR contains a comma *
    ' *           separated list, or both, then the type is Array, and the  *
    ' *           array's entry type is determined by examining the values  *
    ' *           in the array: if they all convert to a single type then   *
    ' *           the array's type is this type: if they all convert to     *
    ' *           more than one type then the array's type is Variant.      *
    ' *                                                                     *
    ' *         Note that the type may not be omitted when a UDT is         *
    ' *         specified.                                                  *
    ' *                                                                     *
    ' *      *  The inspect(report) method checks the state of the object   *
    ' *         for conformity to rules. Note that inspect does not check   *
    ' *         for user error, but instead for errors resulting from       *
    ' *         software bugs and object misuse.                            *
    ' *                                                                     *
    ' *         In inspect, the report parameter must be a ByRef string, and*
    ' *         it is set to an inspection report.  This method returns True*
    ' *         on success; on failure, it returns False and marks the      *
    ' *         object as not usable.                                       *
    ' *                                                                     *
    ' *         The following inspection rules are in effect:               *
    ' *                                                                     *
    ' *         + The object instance must be usable                        *
    ' *                                                                     *
    ' *         + The variable type object objDope must pass its own        *
    ' *           inspection procedure. It must be Unknown or an array type.*
    ' *           If the dope is Unknown, the objValue must be Nothing and  *
    ' *           the following tests are skipped.                          *
    ' *                                                                     *
    ' *         + objValue must be one of the following:                    *
    ' *                                                                     *
    ' *           - Nothing (when the type of the variable is Unknown or    *
    ' *             Null)                                                   *
    ' *                                                                     *
    ' *           - One of the types that represents, in .Net, a Quick Basic*
    ' *             type (Boolean, Byte, Short, Integer, Single, Double,    *
    ' *             or String (when the type of the variable is scalar)     *
    ' *                                                                     *
    ' *           - A collection; the type must be Array or UDT             *
    ' *                                                                     *
    ' *             If the type is Array, this must be an "orthogonal"      *
    ' *             collection, that contains a "balanced"                  *
    ' *             structure of elements representing an array. Each final *
    ' *             element's type must either match the nonvariant type    *
    ' *             in the variable's variableType, or, when the variable's *
    ' *             variableType is Variant, each final element's type must *
    ' *             be the .Net representation of a QB scalar.              *
    ' *                                                                     *
    ' *             The collection must be orthogonal in that it has to be  *
    ' *             a balanced tree. To be a balanced tree, the collection  *
    ' *             must either contain 0..n scalars or 0..n subcollections.*
    ' *             If it consists of subcollections, then each sub         *
    ' *             collection must be balanced and orthogonal.             *
    ' *                                                                     *
    ' *             If the type is UDT the collection must consist,         *
    ' *             exclusively, of qbVariable objects. Each must be a      *
    ' *             scalar, an array, a variant or a UDT.                   *
    ' *                                                                     *
    ' *           - A variant qbVariable that is either abstract (containing*
    ' *             no value) or that is of scalar or UDT type.             *
    ' *                                                                     *
    ' *             Note: at this writing qbVariable does not support       *
    ' *             variants that contain arrays although fromString syntax *
    ' *             allows their specification. This rule shall be changed  *
    ' *             to allow variants that contain arrays when code is      *
    ' *             added to fully support this feature.                    *
    ' *                                                                     *
    ' *         + The toString serialization of the variable must create a  *
    ' *           clone of the variable when used with fromString; note that*
    ' *           variants, arrays and UDTs aren't subject to this rule     *
    ' *                                                                     *
    ' *         + The empirical dope of the variable must be consistent     *
    ' *           with its recorded type. The empirical dope (the type as   *
    ' *           determined by examination of the value) must be either    *
    ' *           the same as or contained in the type.                     *
    ' *                                                                     *
    ' *           Note that scalars only are subject to this rule.          *
    ' *                                                                     *
    ' *         + If the variable is a Variant, its Variant type must match *
    ' *           the type of its entry as seen in the decorated value when *
    ' *           the variable is serialized using toString.                *
    ' *                                                                     *
    ' *           For example, Variant,Byte:Integer(256) isn't valid.       *
    ' *                                                                     *
    ' *      *  The isaNumber method returns True when the variable is a    *
    ' *         scalar of numeric type including Boolean, or a string that  *
    ' *         is a number.                                                *
    ' *                                                                     *
    ' *      *  The isAnUnsignedInteger method returns True when the        *
    ' *         variable is an UNSIGNED integer of any scalar type. Note    *
    ' *         that the variable can be Boolean (but not True, since this  *
    ' *         converts to a signed integer), a string, or a real number   *
    ' *         type...as long as its syntatical representation as a string *
    ' *         is that of an unsigned integer.                             *
    ' *                                                                     *
    ' *      *  The isClear method returns True when the scalar has a       *
    ' *         default value, or the UDT or array  contains all defaults,  *
    ' *         False otherwise.                                            *
    ' *                                                                     *
    ' *      *  The isomorphicVariable(v) method determines whether the     *
    ' *         variable a is an isomorph of the variable in the instance   *
    ' *         according to the rules of the preceding section.            *
    ' *                                                                     *
    ' *         The overloaded syntax isomorphicVariable(v,s) can be used to*
    ' *         derive the explanation as to why v is (or isn't) isomorphic *
    ' *         to the instance. s is a string passed by reference.         *
    ' *                                                                     *
    ' *         UDTs are never isomorphic.                                  *
    ' *                                                                     *
    ' *      *  The isScalar method returns True when the variable is a     *
    ' *         scalar type, or a variant set to a scalar type.             *
    ' *                                                                     *
    ' *      *  The Shared method mkRandomVariable returns a random variable*
    ' *         with random type and value, as an expression that is        *
    ' *         valid input for the fromString method; see the fromString   *
    ' *         and the toString methods for details.                       *
    ' *                                                                     *
    ' *         The fromString returned will have these randomly-selected   *
    ' *         characteristics:                                            *
    ' *                                                                     *
    ' *         + With 10% probability it will be Unknown, with 10% probab- *
    ' *           ility it will be Null.                                    *
    ' *                                                                     *
    ' *         + With 20% probability the fromString will represent a      *
    ' *           scalar and with equal subprobability this will be any of  *
    ' *           the types, Boolean, Byte, Integer, Long, Single, Double,  *
    ' *           or String.                                                *
    ' *                                                                     *
    ' *         + With 20% probability the fromString will be an array, and *
    ' *           this array will contain a variable that has these type    *
    ' *           probabilities, with one exception: there is a 50% proba-  *
    ' *           bility that the variable will be a variant and NO proba-  *
    ' *           bility that the variable will be an array.                *
    ' *                                                                     *
    ' *         + With 20% probability the fromString will be a UDT, and    *
    ' *           this UDT will contain 1..10 scalars, arrays, variants and *
    ' *           UDTs. Each type will have 25% probability.                *
    ' *                                                                     *
    ' *         + With 20% probability the fromString will be a variant, and*
    ' *           this variant will contain a variable that has these type  *
    ' *           probabilities, with one exception: there is a 70% proba-  *
    ' *           bility that the variable will be a scalar and NO proba-   *
    ' *           bility that the variable will be an variant.              *
    ' *                                                                     *
    ' *      *  The mkUnusable method marks the object as not usable.       *
    ' *                                                                     *
    ' *      *  The Shared method mkVariable(fromString) creates and        *
    ' *         returns a new qbVariable object with the specified type and *
    ' *         value in fromString.                                        *
    ' *                                                                     *
    ' *         Note that fromString may be in the syntax type:value or the *
    ' *         syntax value; but in the latter syntax, when the value is   *
    ' *         not a number, it must be quoted using Visual Basic          *
    ' *         conventions.                                                *
    ' *                                                                     *
    ' *         Also see mkVariableFromValue.                               *
    ' *                                                                     *
    ' *      *  The Shared method mkVariableFromValue(value) creates and    *
    ' *         returns a new qbVariable object with the specified scalar   *
    ' *         value. The value operand may be any .Net scalar value of the*
    ' *         type Boolean, Byte, Short, Integer, Long, Single, Double or *
    ' *         String.                                                     *
    ' *                                                                     *
    ' *         Note that when value is a string it should NOT be quoted.   *
    ' *                                                                     *
    ' *         While mkVariable (above) may be used to make an array by    *
    ' *         explicitly specifying array type and members as in the      *
    ' *         example mkVariable("Array,Integer,0,1:0,1"), note that      *
    ' *         mkVariableFromValue cannot create an array. For example,    *
    ' *         mkVariableFromValue("0,1") will create a string.            *
    ' *                                                                     *
    ' *         This method, also, cannot create a variant: the qbVariable, *
    ' *         instead, will have the narrowest possible scalar type. For  *
    ' *         example, mkVariableFromValue(32768) will create a Long      *
    ' *         integer.                                                    *
    ' *                                                                     *
    ' *         Also see mkVariable.                                        *
    ' *                                                                     *
    ' *      *  The event msgEvent(m, L) provides general information. This *
    ' *         exposes the following parameters:                           *
    ' *                                                                     *
    ' *         + m: a general information message                          *
    ' *         + L: nesting level useful in indenting displays             *
    ' *                                                                     *
    ' *         To obtain the message event declare the object WithEvents   *
    ' *         and write a handler.                                        *
    ' *                                                                     *
    ' *      *  The read-write Name property returns and can be set to an   *
    ' *         object instance name.  Name defaults to qbVariable<nnnn>    *
    ' *         <date> <time> where <nnnn> is a sequence number.            *
    ' *                                                                     *
    ' *      *  The Shared method netValue2QBvariable(n) converts the .Net  *
    ' *         object to a nice qbVariable (I'm losin' it).                *
    ' *                                                                     *
    ' *         + If n is Nothing then the Unknown qbVariable is created    *
    ' *           and returned                                              *
    ' *                                                                     *
    ' *         + If n is a .Net scalar that converts to a qb scalar then   *
    ' *           this qbVariable is returned                               *
    ' *                                                                     *
    ' *      *  The New method creates the object instance.  Note that the  *
    ' *         the variable is created with unknown/undefined type and     *
    ' *         value when new has no operands. New(typeValue) creates the  *
    ' *         variable with the specified type and value if typeValue is  *
    ' *         acceptable to the fromString method.                        *
    ' *                                                                     *
    ' *         For example, the following code creates a two dimensional   *
    ' *         Integer array.                                              *
    ' *                                                                     *
    ' *         Dim objArray As qbVariable.qbVariable                       *
    ' *         .                                                           *
    ' *         .                                                           *
    ' *         .                                                           *
    ' *         Try                                                         *
    ' *             objVariable = New qbVariable.qbVariable                 *
    ' *                           ("Array,Integer,0,1,1,3:(1,2,3),(4,5,6)") *
    ' *         Catch...End Try                                             *
    ' *                                                                     *
    ' *      *  The object2XML method returns the state of the object       *
    ' *         formatted using eXtendedMarkupLanguage.                     *
    ' *                                                                     *
    ' *      *  The event progressEvent(a,o,n, c, L, r) and the Shared event*
    ' *         progressEventShared(a,o,n, c, L, r) indicate progress       *
    ' *         through a loop. The progress events expose the following    *
    ' *         parameters:                                                 *
    ' *                                                                     *
    ' *         + a: description of the activity                            *
    ' *         + o: description of the entity being processed              *
    ' *         + n: current entity number                                  *
    ' *         + c: total number of entities                               *
    ' *         + L: the loop's nesting level (from 0)                      *
    ' *         + r: extra comments (remarks)                               *
    ' *                                                                     *
    ' *         To obtain the progress events declare the object WithEvents *
    ' *         and write a handler (probably, sharing common logic) for    *
    ' *         the Shared and state-ful events. Note that state-ful        *
    ' *         instances may emit Shared events.                           *
    ' *                                                                     *
    ' *      *  The stringIdentical(o) method will return True when the     *
    ' *         value or values of the instance are identical to the        *
    ' *         value(s) of the qbVariable object o, False otherwise. The   *
    ' *         instance and o are converted to string format by the        *
    ' *         toString method and the type information (as well as the    *
    ' *         colon separator) is removed for the comparision.            *
    ' *                                                                     *
    ' *      *  The read-write Tag property may be set to any object to     *
    ' *         associate user data with the object instance. Note that     *
    ' *         the associated object, when it is a reference object, is    *
    ' *         not destroyed when the qbVariableType instance is disposed. *
    ' *                                                                     *
    ' *         Also, note that once the variable is Tagged, it will retain *
    ' *         the Tag even if the variable is reassigned a new type and   *
    ' *         value, or reset: to clear the Tag, set it to Nothing.       *
    ' *                                                                     *
    ' *      *  The test(r) method runs tests on the object, and returns    *
    ' *         True to indicate success or False, to indicate failure.     *
    ' *                                                                     *
    ' *         r should be a String which is passed By Reference, and r is *
    ' *         set to a test report on success or failure.                 *
    ' *                                                                     *
    ' *         The following tests are performed:                          *
    ' *                                                                     *
    ' *         + The test variables mentioned in this documentation, as    *
    ' *           well as in the book, are created and inspected for        *
    ' *           validity                                                  *
    ' *                                                                     *
    ' *         + Test cases for variable containment and isomorphism are   *
    ' *           performed (n.b. 9-8: this test is not actually made: cf.  *
    ' *           open issues section below)                                *
    ' *                                                                     *
    ' *      *  The toDescription method returns a description of the value *
    ' *         consisting of the description of its type, followed by      *
    ' *         either "is empty" or "contains nondefault values".          *
    ' *                                                                     *
    ' *      *  For a two-dimensional matrix, only, the toMatrix method     *
    ' *         returns a multi-line, multi-column display of the array     *
    ' *         indexes and values suitable for display in a monospace font.* 
    ' *                                                                     *
    ' *      *  The toString method returns the type and value of the       *
    ' *         qbVariable, in the format described above for fromString.   *
    ' *                                                                     *
    ' *         If the variable is an array, the representation returned    *
    ' *         will be "packed", condensing series of identical elements   *
    ' *         using parenthesized repetition counts as described under    *
    ' *         fromString. The representation will also return default     *
    ' *         values as asterisks.                                        *
    ' *                                                                     *
    ' *         If the scalar value contains the default value appropriate  *
    ' *         to its type, or each member in an array value contains the  *
    ' *         default, toString returns the asterisk.                     *
    ' *                                                                     *
    ' *         Variables are decorated (returned using type(value) syntax) *
    ' *         when the variable type is either variant or array of        *
    ' *         variants.                                                   *
    ' *                                                                     *
    ' *      *  The toStringTypeOnly method returns a string containing     *
    ' *         the type information about the array exclusively.           *
    ' *                                                                     *
    ' *      *  The toStringWithType method returns a string containing     *
    ' *         the type of the variable in addition to its value.                                                      *
    ' *                                                                     *
    ' *      *  The indexed, read-only property UDTmember(x) returns the    *
    ' *         qbVariable object that corresponds to a member of a user    *
    ' *         data type. x may be an index between 1 and the value of     *
    ' *         Dope.UDTmemberCount for the object or it may be a member    *
    ' *         name.                                                       *
    ' *                                                                     *
    ' *      *  The read-only Usable property returns True when the object  *
    ' *         is usable, and False when it is unusable.                   *
    ' *                                                                     *
    ' *      *  The value method returns the .Net value of the variable.    *
    ' *         This method exposes the following overloads:                *
    ' *                                                                     *
    ' *         + value by itself returns the .Net value of an Unknown,     *
    ' *           Null, scalar qbVariable, or a variant containing an       *
    ' *           unknown, Null or scalar qbVariable. For Unknown or Null,  *
    ' *           Nothing is returned: for a scalar, the .Net value of the  *
    ' *           qbVariable is returned, using the narrowest possible .Net *
    ' *           type.                                                     *     
    ' *                                                                     *
    ' *         + value(i) returns the indexed element of a one-dimension   *
    ' *           array using the above rules.                              *
    ' *                                                                     *
    ' *         + value(i,j) returns the indexed element of a two-dimension *
    ' *           array using the above rules.                              *
    ' *                                                                     *
    ' *         + value(s) returns the indexed element of an array using    *
    ' *           the above rules. The ByVal string s should be the comma-  *
    ' *           delimited list of array subscripts.                       *
    ' *                                                                     *
    ' *         + value(member) returns a UDT element where member is its   *
    ' *           member name; note that this syntax can be used to return  *
    ' *           the value of a UDT that is a nested member by using the   *
    ' *           period to separate member names. For example, if the udt  *
    ' *           contains the udt typMember01, and this contains the       *
    ' *           integer intMember02, value("typMember01.intMember02")     *
    ' *           returns the integer member.                               *
    ' *                                                                     *
    ' *      *  The valueSet method changes the .Net value of the variable. *
    ' *         This method exposes the following overloads:                *
    ' *                                                                     *
    ' *         + valueSet(N) assigns the .Net value N to the variable.     *
    ' *           This assignment is strongly typed in that the .Net value  *
    ' *           must convert without error to the pre-existing type in the*
    ' *           object instance.                                          *
    ' *                                                                     *
    ' *           The object type must be scalar or variant with pre-exist- *
    ' *           scalar, Unknown or null contents.                         *
    ' *                                                                     *
    ' *         + valueSet(N,i) assigns the .Net value N to the indexed     *
    ' *           element of a one-dimension array using the above rules.   *
    ' *                                                                     *
    ' *         + valueSet(N,i,j) assigns the .Net value N to the indexed   *
    ' *           element of a two-dimension array using the above rules.   *
    ' *                                                                     *
    ' *         + valueSet(N,s) will assign the .Net value N to the indexed *
    ' *           element of an array using the above rules when the object *
    ' *           is an array (s must be the comma-delimited string of      *
    ' *           array subscripts); valueSet(N,s) will assign N to the UDT *
    ' *           member named by s when the object is a UDT (s may be a    *
    ' *           simple name or a series of nested UDT names): the object  *
    ' *           can have no other type.                                   *
    ' *                                                                     *
    ' *         + valueSet(N,s,t) assigns the .Net value N to the indexed   *
    ' *           element of an array, which is a member of a UDT, using the*
    ' *           above rules.                                              *
    ' *                                                                     *
    ' *         Note: when a string .Net value is assigned to a Boolean     *
    ' *         Quick Basic value it must be True or False (independent of  *
    ' *         its case).                                                  *
    ' *                                                                     *
    ' *      *  The read-write VariableName property returns and can change *
    ' *         the name of the variable. VariableName defaults to typennnn *
    ' *         where:                                                      *
    ' *                                                                     *
    ' *         + type is the three-character Hungarian prefix designating  *
    ' *           the variable type: one of boo, byt, int, lng, sgl, dbl,   *
    ' *           str, vnt, arr, typ, unk or nul                            *
    ' *                                                                     *
    ' *         + For a variant or array the Hungarian prefix vnt or arr is *
    ' *           followed by the Proper case full name of the variant's    *
    ' *           contained type or the array's entry type.                 *
    ' *                                                                     *
    ' *         + nnnn is the variable's sequence number as it appears in   *
    ' *           the Name property                                         *
    ' *                                                                     *
    ' *         Examples of default names include int0001 and               *
    ' *         vntInteger0002.                                             * 
    ' *                                                                     *
    ' *         The name must when assigned conform to Quick Basic rules.   *
    ' *         It must be from 1..31 characters long: it must start with   *
    ' *         a letter: it is restricted to letters, numbers and          *
    ' *         the underscore.                                             *
    ' *                                                                     *
    ' *         See also the Name property and note that the Name denotes   *
    ' *         the object instance while VariableName the variable.        *
    ' *         In particular the VariableName changes when (1) it hasn't   *
    ' *         been assigned except by default and (2) the variable Dope   *
    ' *         (type) is changed, in any way.                              *
    ' *                                                                     *
    ' *                                                                     *
    ' * FROMSTRING BNF SYNTAX --------------------------------------------- *
    ' *                                                                     *
    ' * The lexical syntax of fromString expressions matches that of the    *
    ' * quickBasicEngine itself: blanks can be freely used, and strings are *
    ' * delimited by double quotes (with doubled double quotes representing,*
    ' * inside strings, the occurence of a single double quote.)            *
    ' *                                                                     *
    ' * Note that this object is responsible only for scanning and parsing  *
    ' * the fromStringValue which contains the value(s) of the variable:    *
    ' * parsing of the fromStringType occurs inside the qbVariableType      *
    ' * object.                                                             *
    ' *                                                                     *
    ' *      fromString := fromStringType                                   *
    ' *      fromString := fromStringValue                                  *
    ' *      fromString := fromStringWithValue                              *
    ' *      fromString := fromStringType COLON fromStringValue             *
    ' *      fromString := COLON fromStringValue                            *
    ' *                                                                     *
    ' *      fromStringType := baseType | udt                               *
    ' *      baseType := simpleType | variantType | arrayType               *
    ' *      simpleType := [VT] typeName                                    *
    ' *      typeName := BOOLEAN|BYTE|INTEGER|LONG|SINGLE|DOUBLE|STRING|    *
    ' *                  UNKNOWN|NULL                                       *
    ' *      variantType := abstractVariantType COMMA varType               *
    ' *      varType := simpleType|(arrayType)                              *
    ' *      arrayType := [VT] ARRAY,arrType,boundList                      *
    ' *      arrType := simpleType | abstractVariantType | parUDT           *
    ' *      parUDT := LEFTPARENTHESIS udt RIGHTPARENTHESIS                 *
    ' *      udt := [VT] UDT,typeList                                       *
    ' *      typeList := parMemberType [ COMMA typeList ]                   *
    ' *      parMemberType := LEFTPAR MEMBERNAME,baseType RIGHTPAR          *
    ' *      abstractVariantType := [VT] VARIANT                            *
    ' *      boundList := boundListEntry | boundListEntry COMMA boundList   *
    ' *      boundListEntry := BOUNDINTEGER,BOUNDINTEGER                    *
    ' *      simpleType := [VT] typeName                                    *
    ' *      typeName := BOOLEAN|BYTE|INTEGER|LONG|SINGLE|DOUBLE|STRING|    *
    ' *                  UNKNOWN|NULL                                       *
    ' *      variantType := abstractVariantType,varType                     *
    ' *      varType := simpleType|(arrayType)                              *
    ' *      arrayType := [VT] ARRAY,arrType,boundList                      *
    ' *      arrType := simpleType|abstractVariantType                      *
    ' *      abstractVariantType := [VT] VARIANT                            *
    ' *      boundList := boundListEntry | boundListEntry, boundList        *
    ' *      boundListEntry := BOUNDINTEGER,BOUNDINTEGER                    *
    ' *      fromStringValue := ASTERISK | fromStringNondefault             *
    ' *      fromStringNondefault := arraySlice [ COMMA fromStringValue ] * *
    ' *      arraySlice := elementExpression | ( fromStringNondefault )     *
    ' *      elementExpression := element [ repeater ]                      *
    ' *      element := scalar | decoValue                                  *
    ' *      scalar := NUMBER | VBQUOTEDSTRING | ASTERISK | TRUE | FALSE    *
    ' *      decoValue := quickBasicDecoValue | netDecoValue                *
    ' *      quickBasicDecoValue := QUICKBASICTYPE ( scalar )               *
    ' *      netDecoValue := netDecoValue := [ SYSTEM PERIOD ] IDENTIFIER   *
    ' *                      LEFTPARENTHESIS ANYTHING RIGHTPARENTHESIS      *
    ' *      repeater := LEFTPAR ( INTEGER | ASTERISK ) RIGHTPAR            *
    ' *                                                                     *
    ' *                                                                     *
    ' * MULTITHREADING CONSIDERATIONS ------------------------------------- *
    ' *                                                                     *
    ' * Multiple instances of this class may simultaneously execute in      *
    ' * different threads. However, the same instance may not run code in   *
    ' * multiple threads.                                                   *
    ' *                                                                     *
    ' *                                                                     *
    ' ***********************************************************************

    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJvariableType As qbVariableType.qbVariableType
    Private Shared _INTsequence As Integer
    Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

    ' ***** State *****
    Private Structure TYPstate
        Dim booUsable As Boolean                           ' Object usability
        Dim strName As String                              ' Instance name
        Dim strVariableName As String                      ' Variable name
        Dim booVariableNameDefaults As Boolean             ' True: variable name has default value
                                                           ' False: variable name has been changed
        Dim objDope As qbVariableType.qbVariableType       ' Variable type
        Dim objValue As Object                             ' Variable value: .Net scalar or collection
        Dim objTag As Object                               ' User object  
    End Structure
    Private USRstate As TYPstate

    ' ***** Constants *****
    ' --- Common displayable characters
    Private Const GRAPHIC_CHARACTERS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                                 "abcdefghijklmnopqrstuvwxyz" & _
                                                 "~`!@#$%^&*()_+-={}|[]\:"";'<>,.?/ "
    ' --- Name of the class
    Private Const CLASS_NAME As String = "qbVariable"
    ' --- Inspection rules
    Private Const INSPECTION_USABLE As String = _
    "Object instance must be usable"
    Private Const INSPECTION_DOPE As String = _
    "The variable type object objDope must pass its own " & _
    "inspection procedure. It must be Unknown or an array type. " & _
    "If the dope is Unknown, the objValue must be Nothing and " & _
    "the following tests are skipped."
    Private Const INSPECTION_VALUE As String = _
    "objValue must be one of the following:" & vbNewLine & vbNewLine & _
    "     *  Nothing (when the type of the variable is Unknown or Null)" & vbNewLine & vbNewLine & _
    "     *  One of the types that represents, in .Net, a Quick Basic" & vbNewLine & _
    "        type (Boolean, Byte, Short, Integer, Single, Double," & vbNewLine & _
    "        or String (when the type of the variable is scalar)" & vbNewLine & vbNewLine & _
    "     *  An ""orthogonal"" collection, that contains a ""balanced""" & vbNewLine & _
    "        structure of elements representing an array. Each final" & vbNewLine & _
    "        element's type must either match the nonvariant type" & vbNewLine & _
    "        in the variable's variableType, or, when the variable's" & vbNewLine & _
    "        element's type must either match the nonvariant type" & vbNewLine & _
    "        variableType is Variant, each final element's type must" & vbNewLine & _
    "        be the .Net representation of a QB scalar." & vbNewLine & vbNewLine & _
    "        The collection must be orthogonal in that it has to be" & vbNewLine & _
    "        a balanced tree. To be a balanced tree, the collection" & vbNewLine & _
    "        must either contain 0..n scalars or 0..n subcollections." & vbNewLine & _
    "        If it consists of subcollections, then each sub" & vbNewLine & _
    "        collection must be balanced and orthogonal." & vbNewLine & vbNewLine & _
    "     *  A collection consisting of qbVariables." & vbNewLine & vbNewLine & _
    "     *  A variant that is either abstract or of scalar type."
    Private Const INSPECTION_CLONE As String = _
    "The toString serialization of the nonvariant variable must create a clone " & _
    "of the variable when used with fromString"
    Private Const INSPECTION_EMPIRICALVALUE As String = _
    "The empirical dope derived from the variable must be consistent with " & _
    "the type of the variable. The empirical dope must be identical-to, or contained-in, " & _
    "the original type."
    Private Const INSPECTION_VARIANTCONSISTENCY As String = _
    "If the variable is a Variant, its Variant type must match " & _
    "the type of its entry as seen in the decorated value when " & _
    "the variable is serialized using toString. " & _
    "For example, Variant,Byte:Integer(256) isn't valid."
    ' --- About information Easter egg
    Private Const ABOUTINFO As String = _
        "This class represents the type and value of a Quick Basic variable" & _
        vbNewLine & vbNewLine & _
        "This class was developed commencing on June 24 2003 by" & _
        vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    ' --- Testing
    ' Examples for testing: fromString and expected toString
    '
    ' Note 2-10-2004: no UDT test cases are included. Cf. qbVariable\Archive\
    ' 02102004\udtTestCases.TXT for these cases. At this writing they cause
    ' test failure with an unusable object
    '
    '
    Private Const TEST_EXAMPLES As String = _
    "Integer:4" & vbNewLine & _
    "Integer:vtInteger(4)" & vbNewLine & vbNewLine & _
    "Variant,Integer:4" & vbNewLine & _
    "Variant,Integer:vtInteger(4)" & vbNewLine & vbNewLine & _
    "Array,Integer,0,3:1,2,3,4" & vbNewLine & _
    "Array,Integer,0,3:vtInteger(1),vtInteger(2),vtInteger(3),vtInteger(4)" & vbNewLine & vbNewLine & _
    "Array,Integer,1,2,1,2:(1,2),(1,2)" & vbNewLine & _
    "Array,Integer,1,2,1,2:(vtInteger(1),vtInteger(2)),(vtInteger(1),vtInteger(2))" & vbNewLine & vbNewLine & _
    "Array,Variant,1,2:Integer(1),Long(2)" & vbNewLine & _
    "Array,Variant,1,2:vtInteger(1),vtLong(2)" & vbNewLine & vbNewLine & _
    "Array,Variant,1,1:String(""A"")" & vbNewLine & _
    "Array,Variant,1,1:vtString(""A"")" & vbNewLine & vbNewLine & _
    "Array,Variant,1,2:Integer(1),String(""A"")" & vbNewLine & _
    "Array,Variant,1,2:vtInteger(1),vtString(""A"")" & vbNewLine & vbNewLine & _
    "32768" & vbNewLine & _
    "Long:vtLong(32768)" & vbNewLine & vbNewLine & _
    ":32767" & vbNewLine & _
    "Integer:vtInteger(32767)" & vbNewLine & vbNewLine & _
    "4" & vbNewLine & _
    "Byte:vtByte(4)" & vbNewLine & vbNewLine & _
    "Array,Byte,0,1" & vbNewLine & _
    "Array,Byte,0,1:*" & vbNewLine & vbNewLine & _
    "Array,String,1,10:*" & vbNewLine & _
    "Array,String,1,10:*" & vbNewLine & vbNewLine & _
    "Array,Byte,0,1:*,1" & vbNewLine & _
    "Array,Byte,0,1:vtByte(0),vtByte(1)" & vbNewLine & vbNewLine & _
    "Array,String,1,10:""A""(10)" & vbNewLine & _
    "Array,String,1,10:vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A""),vtString(""A"")" & vbNewLine & vbNewLine & _
    "Array,Variant,0,1,1,2:(32767,""B""),(32768,1)" & vbNewLine & _
    "Array,Variant,0,1,1,2:(vtInteger(32767),vtString(""B"")),(vtLong(32768),vtByte(1))" & _
    vbNewLine & vbNewLine & _
    "Array,Integer,1,2,1,2:(1,2),(3,4)" & vbNewLine & _
    "Array,Integer,1,2,1,2:(vtInteger(1),vtInteger(2)),(vtInteger(3),vtInteger(4))"
    ' Test containment and isomorphism: toString1, toString2, and expected relationship
    '
    ' Test containment not tested and the test cases need to be created
    '
    Private Const TEST_CONTAINMENT As String = _
    ""

    ' ***** Events ****************************************************

    ' --- General information 
    Public Event msgEvent(ByVal strMsg As String, _
                          ByVal intLevel As Integer)

    ' --- Progress reporting 
    Public Event progressEvent(ByVal strActivity As String, _
                               ByVal strEntity As String, _
                               ByVal intEntityNumber As Integer, _
                               ByVal intEntityCount As Integer, _
                               ByVal intLevel As Integer, _
                               ByVal strComments As String)
    Public Shared Event progressEventShared(ByVal strActivity As String, _
                                            ByVal strEntity As String, _
                                            ByVal intEntityNumber As Integer, _
                                            ByVal intEntityCount As Integer, _
                                            ByVal intLevel As Integer, _
                                            ByVal strComments As String)

    ' ***** Object constructor ****************************************

    ' --- No fromString defined
    Public Sub New()
        new_("Unknown")
    End Sub
    ' --- Fromstring defined 
    Public Sub New(ByVal strFromString As String)
        new_(strFromString)
    End Sub
    ' --- Common functionality
    Private Sub new_(ByVal strFromString As String)
        With USRstate
            .strName = ClassName & _
                       sequenceNumber_() & " " & _
                       Now
            Try
                OBJcollectionUtilities = New collectionUtilities.collectionUtilities
                .objDope = New qbVariableType.qbVariableType
            Catch
                errorHandler_("Can't create reference objects: " & _
                              Err.Number & " " & Err.Description, _
                              "new", _
                              "Object is not usable")
                Return
            End Try
            .booVariableNameDefaults = True
            .strVariableName = mkVariableName_(.objDope)
            .booUsable = True
            .booUsable = fromString_(strFromString, "")
            If Not .booUsable Then Return
            .strVariableName = mkVariableName_(.objDope)
        End With
    End Sub

    ' ***** Public procedures **********************************************
    ' *                                                                    *
    ' * Also contains Private procedures, that do work on behalf of a      *
    ' * single Public procedure.                                           *
    ' *                                                                    *
    ' **********************************************************************

    ' ----------------------------------------------------------------------
    ' Return about info
    '
    '
    Public Shared ReadOnly Property About() As String
        Get
            Return (ClassName & vbNewLine & vbNewLine & ABOUTINFO)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Convert class information to XML
    '
    '
    Public Shared Function class2XML() As String
        Return (_OBJutilities.objectInfo2XML(ClassName, _
                                            _OBJutilities.soft2HardParagraph(About), _
                                            True, True))
    End Function

    ' ----------------------------------------------------------------------
    ' Return class name
    '
    '
    Public Shared ReadOnly Property ClassName() As String
        Get
            Return (CLASS_NAME)
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Clear the variable
    '
    '
    Public Function clearVariable() As Boolean
        If Not checkUsable_("clearVariable", "Variable cannot be cleared") Then
            Return (False)
        End If
        With USRstate
            If .objDope.isScalar Then
                .objValue = _OBJvariableType.defaultValue(.objDope.VariableType)
            ElseIf .objDope.isArray Then
                clearArrayCollectionValues_(CType(.objValue, Collection), _
                                            .objDope.defaultValue, _
                                            0)
            ElseIf .objDope.isVariant Then
                If Not (.objValue Is Nothing) Then
                    If (TypeOf .objValue Is qbVariable) Then
                        CType(.objValue, qbVariable).clearVariable()
                    Else
                        .objValue = Nothing
                    End If
                End If
            ElseIf .objDope.isUDT Then
                clearUDTcollectionValues_(CType(.objValue, Collection))
            Else
                .objValue = Nothing
            End If
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Cloner
    '
    '
    Public Function clone() As qbVariable
        If Not checkUsable_("clone", "Returning Nothing") Then
            Return (Nothing)
        End If
        Return (CType(clone_(), qbVariable))
    End Function
    Private Function clone_() As Object Implements ICloneable.Clone
        Dim objClone As qbVariable
        Try
            objClone = New qbVariable
        Catch
            errorHandler_("Cannot create clone object: " & _
                            Err.Number & " " & Err.Description, _
                            "clone_", _
                            "Returning Nothing")
            Return (Nothing)
        End Try
        If Not objClone.fromString(Me.ToString) Then
            errorHandler_("Cannot populate clone object: " & _
                            Err.Number & " " & Err.Description, _
                            "clone_", _
                            "Returning Nothing")
            objClone.dispose()
            Return (Nothing)
        End If
        Return (objClone)
    End Function

    ' ----------------------------------------------------------------------
    ' Compare to another object
    '
    '
    Public Overloads Function compareTo(ByVal objQBvariable2 As qbVariable) As Boolean
        If Not checkUsable_("compareTo", "Returning False") Then Return (False)
        If Not objQBvariable2.Usable Then
            errorHandler_("Compared object isn't Usable", _
                          "compareTo", _
                          "Returning False")
            Return (False)
        End If
        Return (compareTo_(objQBvariable2) <> 0)
    End Function
    Private Overloads Function compareTo_(ByVal objQBvariable2 As Object) As Integer _
           Implements IComparable.CompareTo
        If Not (TypeOf objQBvariable2 Is qbVariable) Then Return (0)
        Return (CInt(IIf(Me.ToString = CType(objQBvariable2, qbVariable).ToString, -1, 0)))
    End Function

    ' ----------------------------------------------------------------------
    ' Test containment
    '
    '
    ' --- No explanation
    Public Overloads Function containedVariable(ByVal objVariable2 As qbVariable) _
           As Boolean
        Dim strToss As String
        Return (containedVariable(objVariable2, strToss))
    End Function
    ' --- Explanation provided
    Public Overloads Function containedVariable(ByVal objVariable2 As qbVariable, _
                                                ByRef strExplanation As String) _
           As Boolean
        If Not checkUsable_("containedVariable", "Returning False") Then
            Return (False)
        End If
        With Me
            Return (Me.Dope.containedTypeWithState(objVariable2.Dope, _
                                                  .Dope, _
                                                  strExplanation))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Dereference the member name
    '
    '
    Public Function derefMemberName(ByVal strName As String) As qbVariable
        If Not checkUsable_("derefMemberName", "Returning Nothing") Then
            Return (Nothing)
        End If
        Dim objInner As qbVariable
        Dim strName1 As String = _OBJutilities.item(strName, 1, ".", False)
        Try
            objInner = Me.UDTmember(strName1)
        Catch : End Try
        If (objInner Is Nothing) Then Return (Nothing)
        If _OBJutilities.items(strName, ".", False) = 1 Then Return (objInner)
        If Not objInner.Dope.isUDT Then
            errorHandler_("Member validly identified by " & strName1 & " " & _
                          "is not a user data type, therefore the additional " & _
                          "names are not valid", _
                          "derefMemberName", _
                          "Returning Nothing")
        End If
        Return (objInner.derefMemberName(_OBJutilities.itemPhrase(strName, 2, -1, ".", False)))
    End Function

    ' ----------------------------------------------------------------------
    ' Disposer
    '
    '
    ' --- Conduct an inspection
    Public Sub dispose() Implements IDisposable.Dispose
        dispose_(False)
    End Sub
    ' --- Exposes inspection option
    Public Function disposeInspect() As Boolean
        dispose_(True)
    End Function
    ' --- Common Logic
    Private Function dispose_(ByVal booInspect As Boolean) As Boolean
        With USRstate
            If Me.Usable AndAlso booInspect Then inspection_()
            .booUsable = False
            If Not (.objDope Is Nothing) Then
                If booInspect Then
                    .objDope.disposeInspect()
                Else
                    .objDope.dispose()
                End If
                .objDope = Nothing
            End If
            If (TypeOf .objValue Is Collection) Then
                Return (OBJcollectionUtilities.collectionClear(CType(.objValue, Collection)))
            End If
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return and change variable type
    '
    '
    Public Property Dope() As qbVariableType.qbVariableType
        Get
            If Not checkUsable_("Dope Get", "Returning Nothing") Then
                Return (Nothing)
            End If
            Return (USRstate.objDope)
        End Get
        Set(ByVal objNewValue As qbVariableType.qbVariableType)
            If Not checkUsable_("Dope Get", "No change made to object") Then
                Return
            End If
            With USRstate
                Dim booIsomorphicDope As Boolean = _
                    isomorphicDope_(.objDope, objNewValue)
                .objDope = objNewValue
                If booIsomorphicDope Then
                    Me.clearVariable()
                ElseIf .objDope.VariableType _
                       <> _
                       qbVariableType.qbVariableType.ENUvarType.vtUnknown Then
                    USRstate.objValue = mkVariableValue_(.objDope)
                End If
                If .booVariableNameDefaults Then
                    .strVariableName = mkVariableName_(.objDope)
                End If
            End With
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Reconstruct dope exclusively from data
    '
    '
    Public Function empiricalDope() As String
        If Not checkUsable_("empiricalDope", "Returning a null string") Then
            Return ("")
        End If
        Return (empiricalDope_(USRstate.objValue).ToString)
    End Function

    ' -----------------------------------------------------------------------
    ' Sets the variable from a string representation  
    '
    '
    Public Function fromString(ByVal strFromstring As String) As Boolean
        If Not checkUsable_("fromString", "No change made to object") Then
            Return (False)
        End If
        Dim strExistingDope As String = Me.Dope.ToString
        If Not Me.resetVariable Then Return (False)
        If Not fromString_(strFromstring, strExistingDope) Then Return (False)
        If USRstate.booVariableNameDefaults Then
            Me.VariableName = mkVariableName_(Me.Dope)
            USRstate.booVariableNameDefaults = True
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Translate the fromString on behalf
    '
    '      fromString := fromStringType                                    
    '      fromString := fromStringValue                                   
    '      fromString := fromStringType COLON fromStringValue             
    '      fromString := COLON fromStringValue             
    '
    '
    Private Function fromString_(ByVal strFromstring As String, _
                                 ByVal strExistingDope As String) _
            As Boolean
        ' --- Add pre-existing dope if syntax is :value
        Dim strFromStringWork As String = Trim(strFromstring)
        If Mid(strFromStringWork, 1, 1) = ":" Then
            strFromStringWork = CStr(IIf(strExistingDope = "", _
                                         Me.Dope.ToString, _
                                         strExistingDope)) & _
                                ":" & _
                                Mid(strFromStringWork, 2)
        End If
        ' --- Scan input expression
        Dim objScanner As qbScanner.qbScanner
        Try
            objScanner = New qbScanner.qbScanner
            With objScanner
                .SourceCode = strFromStringWork
                .scan()
                If .TokenCount < 1 AndAlso Trim(strFromStringWork) <> "" _
                   OrElse _
                   .tokenEndIndex(.TokenCount) < Len(strFromStringWork) Then
                    errorHandler_("Unable to scan the fromString expression: " & _
                                    "unrecognizable characters found", _
                                    "fromString_", _
                                    "Returning False")
                    Return False
                End If
            End With
        Catch objEx As Exception
            errorHandler_("Unable to scan the fromString expression: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString_", _
                          "Returning Nothing")
            Return False
        End Try
        With objScanner
            ' --- Parse the type and the value
            Dim intIndex1 As Integer = 1
            Dim strFromStringType As String
            For intIndex1 = 1 To .TokenCount
                If .tokenTypeAsString(intIndex1) = "tokenTypeColon" Then Exit For
                strFromStringType &= .sourceMid(intIndex1)
            Next intIndex1
            If fromString__parse_fromStringType_(strFromStringType, USRstate.objDope) Then
                If USRstate.objDope.Abstract Then
                    fromString__parse_errorHandler_(strFromstring, _
                                                    "Variant has no type in " & _
                                                    _OBJutilities.enquote(strFromStringType), _
                                                    1)
                End If
                USRstate.objValue = mkVariableValue_(USRstate.objDope)
                If .checkToken(intIndex1, ":") Then
                    If Not fromString__parse_fromStringValue_(objScanner, _
                                                              intIndex1, _
                                                              USRstate.objValue, _
                                                              .TokenCount, _
                                                              0, _
                                                              USRstate.objDope.Dimensions, _
                                                              USRstate.objDope) Then
                        .dispose()
                        Return (False)
                    End If
                End If
            Else
                intIndex1 = 1
                If Not fromString__parse_fromStringValue_(objScanner, _
                                                        intIndex1, _
                                                        USRstate.objValue, _
                                                        .TokenCount, _
                                                        0, -1, _
                                                        USRstate.objDope) Then
                    .dispose()
                    Return (False)
                End If
            End If
            Dim booReturn As Boolean = (intIndex1 >= .TokenCount + 1)
            .dispose()
            Return (booReturn)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' arraySlice := elementExpression | ( fromStringNondefault ) 
    '
    '
    Private Function fromString__parse_arraySlice_(ByVal objScanner As qbScanner.qbScanner, _
                                                   ByRef intIndex As Integer, _
                                                   ByRef objNew As Object, _
                                                   ByVal intEndIndex As Integer, _
                                                   ByVal intParseLevel As Integer, _
                                                   ByRef intDimension As Integer, _
                                                   ByVal intElementIndex As Integer, _
                                                   ByRef objExpectedType _
                                                         As qbVariableType.qbVariableType, _
                                                   ByVal objExpectedTypeParent _
                                                         As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing arraySlice")
        Dim intIndex1 As Integer = intIndex
        If fromString__parse_elementExpression_(objScanner, _
                                                intIndex, _
                                                objNew, _
                                                intEndIndex, _
                                                intParseLevel + 1, _
                                                intDimension, _
                                                intElementIndex, _
                                                objExpectedType) Then
            Return (True)
        End If
        With objScanner
            If Not .checkToken(intIndex, "(", intEndIndex:=intEndIndex) Then
                intIndex = intIndex1 : Return (False)
            End If
            If objExpectedTypeParent.isArray Then intDimension -= 1
            If Not (TypeOf objNew Is Collection) Then
                ' We know that the output object is a container collection
                If Not fromString__parse_promoteCollection_(objNew) Then
                    If objExpectedTypeParent.isArray Then intDimension += 1
                    Return (False)
                End If
            End If
            Dim intEndIndex2 As Integer = .findRightParenthesis(intIndex, intEndIndex)
            If intEndIndex2 > intEndIndex Then
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Unbalanced left parenthesis", _
                                                intIndex)
                Return (False)
            End If
            If Not fromString__parse_fromStringNondefault_(objScanner, _
                                                            intIndex, _
                                                            objNew, _
                                                            intEndIndex2 - 1, _
                                                            intParseLevel + 1, _
                                                            intDimension, _
                                                            objExpectedType) Then
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Left parenthesis is not followed by " & _
                                                "an inner fromString expression", _
                                                intIndex)
                intIndex = intIndex1
                If objExpectedTypeParent.isArray Then intDimension += 1
                Return (False)
            End If
            If intIndex <> intEndIndex2 Then
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Right parenthesis is not preceded by " & _
                                                "an inner fromString expression", _
                                                intIndex)
                intIndex = intIndex1
                If objExpectedTypeParent.isArray Then intDimension += 1
                Return (False)
            End If
            intIndex += 1
            If objExpectedTypeParent.isArray Then intDimension += 1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' decoValue := quickBasicDecoValue | netDecoValue
    '
    '
    Private Function fromString__parse_decoValue_(ByVal objScanner As qbScanner.qbScanner, _
                                                  ByRef intIndex As Integer, _
                                                  ByRef objNew As Object, _
                                                  ByVal intEndIndex As Integer, _
                                                  ByRef objValue As Object, _
                                                  ByVal intParseLevel As Integer, _
                                                  ByRef intDimension As Integer, _
                                                  ByRef objExpectedType _
                                                        As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing decoValue")
        Dim intIndex1 As Integer = intIndex
        If fromString__parse_quickBasicDecoValue_(objScanner, _
                                                    intIndex, _
                                                    objNew, _
                                                    intEndIndex, _
                                                    objValue, _
                                                    intParseLevel + 1, _
                                                    intDimension, _
                                                    objExpectedType) Then
            Return (True)
        End If
        intIndex = intIndex1
        If fromString__parse_netDecoValue_(objScanner, _
                                            intIndex, _
                                            objNew, _
                                            intEndIndex, _
                                            objValue, _
                                            intParseLevel + 1, _
                                            intDimension, _
                                            objExpectedType) Then
            Return (True)
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' element := scalar | decoValue | ASTERISK
    '
    '
    Private Function fromString__parse_element_(ByVal objScanner As qbScanner.qbScanner, _
                                                ByRef intIndex As Integer, _
                                                ByRef objNew As Object, _
                                                ByVal intEndIndex As Integer, _
                                                ByVal intParseLevel As Integer, _
                                                ByRef intDimension As Integer, _
                                                ByVal intElementIndex As Integer, _
                                                ByRef objExpectedType _
                                                      As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing element")
        With objScanner
            Dim objElement As Object
            If fromString__parse_scalar_(objScanner, _
                                         intIndex, _
                                         objNew, _
                                         intEndIndex, _
                                         intParseLevel + 1, _
                                         intDimension, _
                                         intElementIndex, _
                                         objExpectedType) Then
                Return (True)
            End If
            Dim objDecoValue As Object
            If fromString__parse_decoValue_(objScanner, _
                                            intIndex, _
                                            objNew, _
                                            intEndIndex, _
                                            objDecoValue, _
                                            intParseLevel + 1, _
                                            intDimension, _
                                            objExpectedType) Then
                fromString__parse_setValue_(objScanner, _
                                            intIndex, _
                                            objNew, _
                                            objDecoValue, _
                                            objExpectedType, _
                                            intDimension, _
                                            intElementIndex)
                Return (True)
            End If
            If .checkToken(intIndex, "*", intEndIndex:=intEndIndex) Then
                objNew = USRstate.objDope.defaultValue
                Return (True)
            End If
            Return (False)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' elementExpression := element [ repeater ]
    '
    '
    Private Function fromString__parse_elementExpression_(ByVal objScanner As qbScanner.qbScanner, _
                                                          ByRef intIndex As Integer, _
                                                          ByRef objNew As Object, _
                                                          ByVal intEndIndex As Integer, _
                                                          ByVal intParseLevel As Integer, _
                                                          ByRef intDimension As Integer, _
                                                          ByVal intElementIndex As Integer, _
                                                          ByRef objExpectedType _
                                                                As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing elementExpression")
        Dim intIndex1 As Integer = intIndex
        If Not fromString__parse_element_(objScanner, _
                                          intIndex, _
                                          objNew, _
                                          intEndIndex, _
                                          intParseLevel + 1, _
                                          intDimension, _
                                          intElementIndex, _
                                          objExpectedType) Then
            Return (False)
        End If
        Dim intCount As Integer
        If fromString__parse_repeater_(objScanner, _
                                       intIndex, _
                                       objNew, _
                                       intEndIndex, _
                                       intCount, _
                                       intParseLevel + 1, _
                                       intDimension, _
                                       objExpectedType) Then
            If Not (TypeOf objNew Is Collection) Then
                fromString__parse_promoteCollection_(objNew)
            End If
            With CType(objNew, Collection)
                Dim intIndex2 As Integer
                Dim intIndex3 As Integer = .Count
                For intIndex2 = 1 To intCount
                    If Not fromString__parse_setValue_(objScanner, _
                                                       intIndex, _
                                                       objNew, _
                                                       .Item(1), _
                                                       objExpectedType, _
                                                       intDimension, _
                                                       intIndex2) Then
                        Exit Function
                    End If
                Next intIndex2
            End With
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Parse errors on behalf of the fromString parser
    '
    '
    Private Sub fromString__parse_errorHandler_(ByVal strFromString As String, _
                                                ByVal strError As String, _
                                                ByVal intIndex As Integer)
        With _OBJutilities
            errorHandler_("Error at token position " & intIndex & " " & _
                          "in the fromString expression " & _
                          .enquote(.ellipsis(strFromString, 32)) & ": " & _
                          strError, _
                          "fromString__parse_", _
                          "No change will be made to the variable object")
        End With
    End Sub

    ' -----------------------------------------------------------------------
    ' fromStringNondefault := arraySlice [ , fromStringValue ] *
    '
    '
    Private Function fromString__parse_fromStringNondefault_(ByVal objScanner As qbScanner.qbScanner, _
                                                             ByRef intIndex As Integer, _
                                                             ByRef objNew As Object, _
                                                             ByVal intEndIndex As Integer, _
                                                             ByVal intParseLevel As Integer, _
                                                             ByRef intDimension As Integer, _
                                                             ByRef objExpectedType _
                                                                   As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing fromStringNondefault")
        Dim intIndex1 As Integer = intIndex
        Dim objNew2 As Object
        If Not objExpectedType.isArray AndAlso Not objExpectedType.isUDT Then
            objNew2 = objNew
        End If
        If Not fromString__parse_arraySlice_(objScanner, _
                                             intIndex, _
                                             objNew2, _
                                             intEndIndex, _
                                             intParseLevel + 1, _
                                             intDimension, _
                                             1, _
                                             memberType_(objExpectedType, 1), _
                                             objExpectedType) Then
            Return (False)
        End If
        Dim intIndex2 As Integer
        Dim objNewRHS As Object
        Dim intElementIndex As Integer = 1
        If objExpectedType.isArray Then
            fromString__parse_promoteArray_(objNew2, intDimension)
        End If
        Do
            If Not objScanner.checkTokenByTypeName(intIndex, _
                                                    "Comma", _
                                                    intEndIndex:=intEndIndex) Then
                Exit Do
            End If
            Dim colHandle As Collection = CType(objNew2, Collection)
            intElementIndex = intElementIndex + 1
            intIndex2 = findBalancedComma_(objScanner, _
                                           intIndex, _
                                           intEndIndex)
            If objExpectedType.isArray Then intDimension -= 1
            If Not fromString__parse_fromStringValue_(objScanner, _
                                                      intIndex, _
                                                      objNewRHS, _
                                                      intIndex2 _
                                                      - _
                                                      1, _
                                                      intParseLevel + 1, _
                                                      intDimension, _
                                                      memberType_(objExpectedType, intElementIndex)) Then
                If objExpectedType.isArray Then intDimension += 1
                intIndex = intIndex1
                Return False
            End If
            If objExpectedType.isArray Then intDimension += 1
            Try
                With colHandle
                    If intElementIndex <= colHandle.Count Then
                        .Remove(intElementIndex)
                    Else
                        If objExpectedType.isArray _
                           AndAlso _
                           intElementIndex - 1 > objExpectedType.UpperBound(intDimension) Then
                            fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                            "Cannot specify elements beyond upper bound of " & _
                                                            objExpectedType.UpperBound(intDimension), _
                                                            intIndex)
                        End If
                    End If
                    .Add(objNewRHS, , intElementIndex)
                End With
            Catch
                errorHandler_("Not able to attach list member " & _
                            _OBJutilities.object2String(objNewRHS) & " " & _
                            "to end of collection: " & _
                            Err.Number & " " & Err.Description, _
                            "fromString__parse_fromStringNondefault_", _
                            "Marking object unusable and returning false")
                Return False
            End Try
            intIndex = intIndex2
        Loop
        objNew = objNew2
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' Parses the fromString type, using the qbVariableType object's parser
    '
    '
    ' Note that there's a flaw in the qbVariableType, because it has its own
    ' scanner and doesn't use the qbScanner as does this object. This means
    ' that this code unnecessarily rebuilds the source code of the type from
    ' the scanned tokens and sends it to qbVariableType, where the source is
    ' rescanned.
    '
    ' This has been noted as an Issue in qbVariableType at this time because
    ' it's only an efficiency consideration.
    '
    '
    Public Function fromString__parse_fromStringType_(ByVal strFromstring As String, _
                                                      ByRef objType As qbVariableType.qbVariableType) _
           As Boolean
        Dim objBackup As qbVariableType.qbVariableType = objType.clone
        Try
            If Not objType.fromString(strFromstring) Then Return (False)
        Catch
            Try
                objType.dispose()
            Catch : End Try
            objType = objBackup.clone
            Return (False)
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' fromStringValue := ASTERISK | fromStringNondefault
    '
    '
    Private Function fromString__parse_fromStringValue_(ByVal objScanner As qbScanner.qbScanner, _
                                                        ByRef intIndex As Integer, _
                                                        ByRef objNew As Object, _
                                                        ByVal intEndIndex As Integer, _
                                                        ByVal intParseLevel As Integer, _
                                                        ByRef intDimension As Integer, _
                                                        ByRef objExpectedType _
                                                              As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing fromStringValue")
        If objScanner.checkToken(intIndex, "*", intEndIndex:=intEndIndex) Then
            If intIndex = intEndIndex + 1 Then Return (True)
            intIndex -= 1
        End If
        Return (fromString__parse_fromStringNondefault_(objScanner, _
                                                       intIndex, _
                                                       objNew, _
                                                       intEndIndex, _
                                                       intParseLevel + 1, _
                                                       intDimension, _
                                                       objExpectedType))
    End Function

    ' -----------------------------------------------------------------------
    ' Make the collection on behalf of the fromString parse
    '
    '
    Private Function fromString__parse_mkCollection_() As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch objEx As Exception
            errorHandler_("Unable to create the collection: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString__parse_fromString_", _
                          "Marking object unusable and returning Nothing")
            Me.mkUnusable() : Return (Nothing)
        End Try
        Return (colNew)
    End Function

    ' -----------------------------------------------------------------------
    ' netDecoValue := [ SYSTEM PERIOD ] IDENTIFIER LEFTPARENTHESIS ANYTHING 
    '                 RIGHTPARENTHESIS  
    '
    '
    Private Function fromString__parse_netDecoValue_ _
                     (ByVal objScanner As qbScanner.qbScanner, _
                      ByRef intIndex As Integer, _
                      ByRef objNew As Object, _
                      ByVal intEndIndex As Integer, _
                      ByRef objValue As Object, _
                      ByVal intParseLevel As Integer, _
                      ByRef intDimension As Integer, _
                      ByRef objExpectedType _
                            As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing netDecoValue")
        With objScanner
            Dim intIndex1 As Integer = intIndex
            If .checkToken(intIndex, _
                               "SYSTEM", _
                               intEndIndex:=intEndIndex) Then
                If Not .checkTokenByTypeName(intIndex, "Period", intEndIndex:=intEndIndex) Then
                    intIndex = intIndex1 : Return (False)
                End If
            End If
            If Not .checkTokenByTypeName(intIndex, "Identifier", intEndIndex:=intEndIndex) Then
                intIndex = intIndex1 : Return (False)
            End If
            If Not .checkToken(intIndex, _
                               "(", _
                               intEndIndex:=intEndIndex) Then
                intIndex = intIndex1 : Return (False)
            End If
            Dim intIndexRP As Integer = .findRightParenthesis(intIndex, _
                                                              intEndIndex:=intEndIndex)
            If intIndexRP > .TokenCount Then
                intIndex = intIndex1 : Return (False)
            End If
            intIndex = intIndexRP
            Try
                objValue = _OBJutilities.string2Object(vtPrefixRemove(.sourceMid(intIndex1)) & _
                                                       .sourceMid(intIndex1 + 1, intIndex - intIndex1))
            Catch
                intIndex = intIndex1 : Return (False)
            End Try
            intIndex += 1
            Return (True)
        End With
    End Function


    ' -----------------------------------------------------------------------
    ' Make the object into a collection of sufficient depth to represent
    ' the array
    '
    '
    Private Function fromString__parse_promoteArray_(ByRef objItem As Object, _
                                                     ByVal intDimension As Integer) _
            As Boolean
        Dim intIndex1 As Integer
        Do
            intIndex1 = collectionDepth_(objItem)
            If intIndex1 = intDimension Then Exit Do
            If Not fromString__parse_promoteCollection_(objItem) Then Return False
        Loop
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' Promote a scalar to a collection or a collection to a collection: in
    ' both cases original item becomes item(1) of 1
    '
    '
    Private Function fromString__parse_promoteCollection_(ByRef objItem As Object) _
            As Boolean
        Dim colNew As Collection = fromString__parse_mkCollection_()
        If colNew Is Nothing Then Return (False)
        Try
            colNew.Add(objItem)
        Catch
            errorHandler_("Unable to initialize collection to scalar: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString__parse_arraySlice_", _
                          "Marking object unusable: returning False")
            Me.mkUnusable() : Return (False)
        End Try
        objItem = colNew
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' quickBasicDecoValue := QUICKBASICTYPE ( scalar )  
    '
    '
    Private Function fromString__parse_quickBasicDecoValue_ _
                     (ByVal objScanner As qbScanner.qbScanner, _
                      ByRef intIndex As Integer, _
                      ByRef objNew As Object, _
                      ByVal intEndIndex As Integer, _
                      ByRef objValue As Object, _
                      ByVal intParseLevel As Integer, _
                      ByRef intDimension As Integer, _
                      ByRef objExpectedType _
                            As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing quickBasicDecoValue")
        With objScanner
            Dim intIndex1 As Integer = intIndex
            If Not .checkTokenByTypeName(intIndex, "Identifier", intEndIndex:=intEndIndex) Then
                intIndex = intIndex1 : Return (False)
            End If
            Dim strType As String = .sourceMid(intIndex - 1)
            If Not _OBJvariableType.isScalarType(strType) Then
                intIndex = intIndex1 : Return (False)
            End If
            If Not .checkToken(intIndex, "(", intEndIndex:=intEndIndex) Then
                intIndex = intIndex1
                Return (False)
            End If
            Dim intEndIndex2 As Integer = .findRightParenthesis(intIndex, intEndIndex)
            If intEndIndex2 > .TokenCount Then
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Unbalanced left parenthesis", _
                                                intIndex)
                intIndex = intIndex1
                Return (False)
            End If
            Dim strDeco As String = _OBJvariableType.qbDomain2NetType _
                                    (_OBJvariableType.string2enuVarType(.sourceMid(intIndex - 2))).ToString & _
                                    "(" & _
                                    .sourceMid(intIndex, intEndIndex2 - intIndex) & _
                                    ")"
            Try
                objValue = _OBJutilities.string2Object(strDeco)
            Catch
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Invalid decorated value " & _
                                                _OBJutilities.enquote(strDeco), _
                                                intIndex)
                Return (False)
            End Try
            intIndex = intEndIndex2 + 1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' repeater := LEFTPAR ( INTEGER | ASTERISK ) RIGHTPAR 
    '
    '
    Private Function fromString__parse_repeater_(ByVal objScanner As qbScanner.qbScanner, _
                                                 ByRef intIndex As Integer, _
                                                 ByRef objNew As Object, _
                                                 ByVal intEndIndex As Integer, _
                                                 ByRef intCount As Integer, _
                                                 ByVal intParseLevel As Integer, _
                                                 ByRef intDimension As Integer, _
                                                 ByRef objExpectedType _
                                                       As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing repeater")
        With objScanner
            Dim intIndex1 As Integer = intIndex
            Dim objScalar As Object
            If Not .checkToken(intIndex, "(", intEndIndex:=intEndIndex) Then
                Return (False)
            End If
            If .checkTokenByTypeName(intIndex, "UnsignedInteger", intEndIndex:=intEndIndex) Then
                intCount = CInt(.sourceMid(intIndex - 1))
            ElseIf .checkToken(intIndex, _
                               "*", _
                               intEndIndex:=intEndIndex) Then
                With USRstate.objDope
                    If intDimension > .Dimensions Then
                        fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                        "Can't use repeat * at a new dimension", _
                                                        objScanner.tokenStartIndex(intIndex - 1))
                        intCount = 0
                        Return (False)
                    End If
                    If Not (TypeOf objNew Is Collection) Then
                        fromString__parse_promoteCollection_(objNew)
                    End If
                    intCount = .UpperBound(intDimension) - CType(objNew, Collection).Count
                End With
            Else
                fromString__parse_errorHandler_(.SourceCode, _
                                                "Invalid repeater", _
                                                .tokenStartIndex(intIndex - 1))
                Return (False)
            End If
            If Not .checkToken(intIndex, ")", intEndIndex:=intEndIndex) Then
                Return (False)
            End If
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' scalar := NUMBER | VBQUOTEDSTRING | TRUE | FALSE | ASTERISK
    '
    '
    Private Function fromString__parse_scalar_(ByVal objScanner As qbScanner.qbScanner, _
                                               ByRef intIndex As Integer, _
                                               ByRef objNew As Object, _
                                               ByVal intEndIndex As Integer, _
                                               ByVal intParseLevel As Integer, _
                                               ByRef intDimension As Integer, _
                                               ByVal intElementIndex As Integer, _
                                               ByRef objExpectedType _
                                                     As qbVariableType.qbVariableType) _
            As Boolean
        RaiseEvent progressEvent("Compiling fromStringValue expression", _
                                        "token", _
                                        intIndex, _
                                        intEndIndex, _
                                        intParseLevel, _
                                        "parsing scalar")
        With objScanner
            Dim intIndex1 As Integer = intIndex
            Dim objScalar As Object
            If objScanner.checkTokenByTypeName(intIndex, _
                                                "String", _
                                                intEndIndex:=intEndIndex) Then
                objScalar = _
                    _OBJutilities.dequote(.sourceMid(intIndex - 1))
            ElseIf intIndex > intEndIndex Then
                Return (False)
            Else
                Dim intSignum As Integer = 1
                If .checkToken(intIndex, "-", intEndIndex:=intEndIndex) Then
                    intSignum = -1
                Else
                    .checkToken(intIndex, "+", intEndIndex:=intEndIndex)
                End If
                If .checkTokenByTypeName(intIndex, _
                                            "UnsignedRealNumber", _
                                            intEndIndex:=intEndIndex) Then
                    objScalar = _OBJvariableType.netValue2QBvalue _
                                (CDbl(.sourceMid(intIndex - 1)) * CDbl(intSignum))
                Else
                    intIndex = intIndex1
                    If .checkToken(intIndex, "True", intEndIndex) Then
                        objScalar = True
                    ElseIf .checkToken(intIndex, "False", intEndIndex) Then
                        objScalar = False
                    ElseIf .checkToken(intIndex, "*", intEndIndex) Then
                        objScalar = objExpectedType.scalarDefault
                    Else
                        Return (False)
                    End If
                End If
            End If
            Return (fromString__parse_setValue_(objScanner, _
                                               intIndex, _
                                               objNew, _
                                               objScalar, _
                                               objExpectedType, _
                                               intDimension, _
                                               intElementIndex))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Assign value by direct assignment or collection addition
    '
    '
    Private Function fromString__parse_setValue_(ByVal objScanner As qbScanner.qbScanner, _
                                                 ByVal intIndex As Integer, _
                                                 ByRef objNew As Object, _
                                                 ByVal objValue As Object, _
                                                 ByRef objExpectedType _
                                                       As qbVariableType.qbVariableType, _
                                                 ByRef intDimension As Integer, _
                                                 ByVal intElementIndex As Integer) _
            As Boolean
        Dim objValueHandle As Object
        If objExpectedType.isVariant Then
            ' When parsing to a variant, we need to make the contained variable
            Dim strType As String = _OBJvariableType.netType2QBdomain(objValue.GetType.ToString).ToString
            If objExpectedType.innerType.isScalar Then
                strType = objExpectedType.innerType.ToString
                objValueHandle = _
                _OBJvariableType.netValue2QBvalue(objValue, _OBJvariableType.vtPrefixAdd(strType))
            Else
                objExpectedType.fromString("Variant," & strType)
                objValueHandle = New qbVariable(strType & _
                                                ":" & _
                                                _OBJutilities.object2String(objValue))
            End If
        ElseIf objExpectedType.isUDT Then
            ' When parsing to a UDT we also need to make the variable
            objValueHandle = New qbVariable(objExpectedType.innerType(intElementIndex).ToString & _
                                            ":" & _
                                            _OBJutilities.object2String(CStr(objValue)))
        Else
            objValueHandle = objValue
        End If
        If (TypeOf objNew Is Collection) Then
            If objExpectedType.isArray AndAlso intElementIndex > objExpectedType.BoundSize(intDimension) _
               OrElse _
               objExpectedType.isUDT AndAlso intElementIndex > objExpectedType.UDTmemberCount Then
                fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                "Too many elements specified for array", _
                                                intIndex)
                Return (False)
            End If
            Try
                Dim colHandle As Collection = CType(objNew, Collection)
                With colHandle
                    If intElementIndex <= .Count Then .Remove(intElementIndex)
                    Dim objQBvalue As Object
                    If (TypeOf objValueHandle Is qbVariable) Then
                        objValueHandle = CType(objValueHandle, qbVariable).value
                    End If
                    If objExpectedType.isVariant Then
                        If objExpectedType.Abstract Then
                            objQBvalue = _OBJvariableType.netValue2QBvalue(objValueHandle)
                        Else
                            objQBvalue = _OBJvariableType.netValue2QBvalue(objValueHandle, _
                                                                           objExpectedType.VarType)
                        End If
                    Else
                        objQBvalue = _OBJvariableType.netValue2QBvalue(objValueHandle, _
                                                                       objExpectedType.VariableType)
                    End If
                    .Add(objQBvalue, , intElementIndex)
                End With
            Catch ex As Exception
                errorHandler_("Cannot add member to collection: " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              "Marking object as unusable: returning False")
                Me.mkUnusable() : Return (False)
            End Try
        Else
            If (TypeOf objValueHandle Is qbVariable) Then
                objNew = objValueHandle
            Else
                Dim booOK As Boolean
                Try
                    booOK = valueSet_(objValueHandle, objExpectedType, objNew)
                Catch : End Try
                If Not booOK Then
                    fromString__parse_errorHandler_(objScanner.SourceCode, _
                                                    "Invalid value type", _
                                                    intIndex)
                    Return (False)
                End If
            End If
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Inspect the variable type
    '
    '
    Public Function inspect(ByRef strReport As String) As Boolean
        Dim booInspection As Boolean = True
        strReport = "Inspection of variable object " & _
                    _OBJutilities.enquote(Me.Name) & " " & _
                    "(" & _OBJutilities.enquote(_OBJutilities.ellipsis(Me.ToString, 64)) & ") " & _
                    "at " & Now & _
                    vbNewLine & vbNewLine & vbNewLine
        With USRstate
            If _OBJutilities.inspectionAppend(strReport, _
                                              INSPECTION_USABLE, _
                                              .booUsable, _
                                              booInspection) Then
                Dim strSubReport As String
                Dim booOK As Boolean
                booOK = .objDope.inspect(strSubReport)
                _OBJutilities.inspectionAppend(strReport, _
                                                  _OBJutilities.string2Box _
                                                  (_OBJutilities.soft2HardParagraph(INSPECTION_DOPE) & _
                                                   vbNewLine & vbNewLine & _
                                                   _OBJutilities.string2Box(strSubReport)), _
                                                  booOK, _
                                                  booInspection, _
                                                  "")
                booOK = inspectValue_(.objValue, .objDope, strSubReport)
                _OBJutilities.inspectionAppend(strReport, _
                                                  _OBJutilities.string2Box _
                                                  (INSPECTION_VALUE & _
                                                   vbNewLine & vbNewLine & _
                                                   _OBJutilities.copies("-", 76) & _
                                                   vbNewLine & vbNewLine & _
                                                   _OBJutilities.soft2HardParagraph(strSubReport)), _
                                                  booOK, _
                                                  booInspection, _
                                                  "")
                If booOK Then
                    booOK = inspectClone_(Me.Dope.isArray, strSubReport)
                    _OBJutilities.inspectionAppend(strReport, _
                                                    _OBJutilities.string2Box _
                                                    (_OBJutilities.soft2HardParagraph(INSPECTION_CLONE) & _
                                                        vbNewLine & vbNewLine & _
                                                        _OBJutilities.copies("-", 76) & _
                                                        vbNewLine & vbNewLine & _
                                                        _OBJutilities.soft2HardParagraph(strSubReport)), _
                                                    booOK, _
                                                    booInspection, _
                                                    "")
                    booOK = booOK AndAlso inspectEmpiricalValue_(strSubReport)
                    _OBJutilities.inspectionAppend(strReport, _
                                                    _OBJutilities.string2Box _
                                                    (_OBJutilities.soft2HardParagraph(INSPECTION_EMPIRICALVALUE) & _
                                                        vbNewLine & vbNewLine & _
                                                        _OBJutilities.copies("-", 76) & _
                                                        vbNewLine & vbNewLine & _
                                                        _OBJutilities.soft2HardParagraph(strSubReport)), _
                                                    booOK, _
                                                    booInspection, _
                                                    "")
                    If Me.Dope.isVariant Then
                        booOK = inspectVariantConsistency_()
                        _OBJutilities.inspectionAppend(strReport, _
                                                        _OBJutilities.string2Box _
                                                        (_OBJutilities.soft2HardParagraph(INSPECTION_VARIANTCONSISTENCY) & _
                                                            vbNewLine & vbNewLine & _
                                                            _OBJutilities.copies("-", 76) & _
                                                            vbNewLine & vbNewLine & _
                                                            _OBJutilities.soft2HardParagraph(strSubReport)), _
                                                        booOK, _
                                                        booInspection, _
                                                        "")
                    End If
                End If
            End If
            If Not booInspection Then Me.mkUnusable()
            Return (booInspection)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Inspect collection (that represents array) on behalf of inspect
    '
    '
    Private Function inspectArrayCollection_(ByVal colArray As Collection, _
                                             ByVal objDope As qbVariableType.qbVariableType, _
                                             ByRef strReport As String) _
            As Boolean
        With colArray
            Dim booCollection As Boolean = (.Count >= 1 AndAlso TypeOf .Item(1) Is Collection)
            Dim intExpectedCount As Integer
            If booCollection Then intExpectedCount = CType(.Item(1), Collection).Count
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                If booCollection Then
                    If intIndex1 > 1 _
                       AndAlso _
                       (Not (TypeOf .Item(intIndex1) Is Collection) _
                        OrElse _
                        CType(.Item(intIndex1), Collection).Count <> intExpectedCount) Then
                        strReport = "At entry " & intIndex1 & ", " & _
                                    "a noncollection was found or a collection with an " & _
                                    "unexpected count: the collection representing the " & _
                                    "array is not orthogonal for this reason"
                        Return (False)
                    End If
                Else
                    Dim objValue As Object = .Item(intIndex1)
                    If (TypeOf objValue Is qbVariable) Then
                        objValue = CType(objValue, qbVariable).value
                    End If
                    If Not _OBJvariableType.containedType _
                            (_OBJvariableType.netValue2QBdomain(objValue), _
                                objDope.VarType) Then
                        strReport = "At entry " & intIndex1 & ", " & _
                                    "an item's type is not valid"
                        Return (False)
                    End If
                End If
            Next intIndex1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Clone and compare, on behalf of inspect
    '
    '
    Private Function inspectClone_(ByVal booArray As Boolean, _
                                   ByRef strReport As String) As Boolean
        strReport = ""
        If Me.Dope.isVariant OrElse Me.Dope.isUDT OrElse Me.Dope.isArray Then
            strReport = "The test is not applicable to Variants, UDTs, or Arrays"
            Return (True)
        End If
        Dim objClone As qbVariable
        Try
            objClone = New qbVariable
        Catch ex As Exception
            errorHandler_("Cannot create qbVariable: " & Err.Number & " " & Err.Description, _
                          "inspectClone_", _
                          "Marking object unusable and returning False")
            Me.mkUnusable() : Return (False)
        End Try
        Dim strTostring As String
        If Not Me.Dope.isNull AndAlso Not Me.Dope.isUnknown Then
            strTostring = ":" & _OBJutilities.itemPhrase(Me.ToString, 2, -1, ":", False)
        End If
        If Not objClone.fromString(Me.Dope.ToString & strTostring) Then Return (False)
        If booArray Then
            Return (Me.isomorphicVariable(objClone, strReport))
        Else
            With objClone
                With .Dope
                    If .isNull OrElse .isUnknown OrElse .isVariant AndAlso .VarType = ENUvarType.vtNull Then
                        With Me.Dope
                            strReport = "This object must clone to a Null, an Unknown, or a null Variant"
                            Return (.isNull OrElse .isUnknown OrElse .isVariant AndAlso .VarType = ENUvarType.vtNull)
                        End With
                    End If
                End With
                strReport = "Instance toString: " & _
                            _OBJutilities.enquote(Me.ToString) & ": " & _
                            "clone toString: " & _
                            _OBJutilities.enquote(.ToString)
                Return (.compareTo(Me))
            End With
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Clone and compare, on behalf of inspect
    '
    '
    Private Function inspectEmpiricalValue_(ByRef strReport As String) As Boolean
        With Me.Dope
            If Not .isScalar Then
                strReport = "This test only applies to scalar variables"
                Return (True)
            End If
        End With
        Try
            Dim booOK As Boolean
            Dim objClone As qbVariableType.qbVariableType = New qbVariableType.qbVariableType
            Dim strExplanation As String
            objClone.fromString(Me.empiricalDope)
            If (USRstate.objDope Is Nothing) Then
                booOK = _OBJvariableType.containedType(objClone, Me.Dope)
            Else
                booOK = USRstate.objDope.containedTypeWithState(objClone, Me.Dope, strExplanation)
            End If
            strReport = "Empirical dope " & _
                        objClone.ToString & " " & _
                        "is " & _
                        CStr(IIf(booOK, "", "NOT ")) & " " & _
                        "contained in actual dope " & _
                        Me.Dope.ToString & _
                        CStr(IIf(strExplanation = "", "", ": " & strExplanation))
            Return (booOK)
        Catch
            strReport = "Error in inspection of empirical value: " & _
                        Err.Number & " " & Err.Description
            Return (False)
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Inspect value on behalf of inspect
    '
    '
    Private Function inspectValue_(ByVal objValue As Object, _
                                   ByVal objDope As qbVariableType.qbVariableType, _
                                   ByRef strReport As String) _
            As Boolean
        With objDope
            Dim booOK As Boolean
            strReport = ""
            If .VariableType = ENUvarType.vtNull OrElse .VariableType = ENUvarType.vtUnknown Then
                booOK = (objValue Is Nothing)
                strReport = "Variable type is Null or Unknown: " & _
                            "the value " & _
                            CStr(IIf(booOK, "is", "isn't")) & " " & _
                            "Nothing"
            ElseIf .isArray Then
                If Not (TypeOf objValue Is Collection) Then
                    strReport = "The value of an array isn't a Collection"
                Else
                    booOK = inspectArrayCollection_(CType(objValue, Collection), _
                                                    objDope, _
                                                    strReport)
                    strReport = "For an array, the value is " & _
                                CStr(IIf(booOK, "a valid", "an invalid")) & " " & _
                                "collection: " & _
                                strReport
                End If
            ElseIf .isUDT Then
                If Not (TypeOf objValue Is Collection) Then
                    strReport = "The value of a UDT isn't a Collection"
                Else
                    booOK = inspectUDTCollection_(CType(objValue, Collection))
                    strReport = "For a user defined type, the value is " & _
                                CStr(IIf(booOK, "a valid", "an invalid")) & " " & _
                                "collection"
                End If
            ElseIf .isScalar Then
                With _OBJvariableType
                    Dim enuType As qbVariableType.qbVariableType.ENUvarType = ENUvarType.vtUnknown
                    Try
                        enuType = .netType2QBdomain(objValue.GetType.ToString)
                    Catch : End Try
                    booOK = .isScalarType(enuType)
                    strReport = "The variableType is scalar " & _
                                CStr(IIf(Not booOK, _
                                         "but the type of the value is not scalar", _
                                         "and the type of the value is scalar"))
                End With
            ElseIf .isVariant Then
                booOK = .Abstract OrElse .isScalarType(.VarType) OrElse .VarType = ENUvarType.vtNull
                strReport = "The variable type is variant " & _
                                CStr(IIf(Not booOK, _
                                         "but the concrete variable doesn't contain scalar, abstract or null type", _
                                         "and the variable is either abstract, null or contains scalar type"))
            Else
                strReport = "The variable type is unrecognizable"
            End If
            Return (booOK)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Inspect variant consistency on behalf of inspect
    '
    '
    Public Function inspectVariantConsistency_() As Boolean
        With _OBJutilities
            Dim strType As String = .item(Me.ToString, 2, ":", False)
            strType = Mid(strType, 1, InStr(strType & "(", "(") - 1)
            Return _OBJvariableType.vtPrefixRemove(.listItem(.item(Me.ToString, 1, ":", False), 2)) _
                   = _
                   _OBJvariableType.vtPrefixRemove(strType)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Test for a numeric scalar
    '
    '
    Public Function isaNumber() As Boolean
        If Not checkUsable_("isaNumber", "Returning False") Then Return (False)
        With Me
            If Not .Dope.isScalar _
               AndAlso _
               Not (.Dope.isVariant OrElse isScalarType(.Dope.VarType)) Then Return (False)
            If .Dope.VariableType = ENUvarType.vtString Then Return (isaNumber_(.value))
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Test for a numeric .Net value
    '
    '
    Private Function isaNumber_(ByVal objValue As Object) As Boolean
        Try
            Dim dblValue As Double = CDbl(objValue)
        Catch
            Return (False)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Test for an unsigned integer scalar
    '
    '
    Public Function isAnUnsignedInteger() As Boolean
        If Not checkUsable_("isAnUnsignedInteger", "Returning False") Then Return (False)
        With Me
            If Not .Dope.isScalar Then Return (False)
            Dim objScanner As qbScanner.qbScanner
            Return (objScanner.isInteger(CStr(Me.value)))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Test for a variable that contains default value(s)
    '
    '
    Public Function isClear() As Boolean
        If Not checkUsable_("isClear", "Returning False") Then Return (False)
        Return (_OBJutilities.item(Me.ToString, 2, ":", False) = "*")
    End Function

    ' ----------------------------------------------------------------------
    ' Test mutual containment
    '
    '
    Public Function isomorphicVariable(ByVal objVariable2 As qbVariable) _
           As Boolean
        Dim strToss As String
        Return (isomorphicVariable(objVariable2, strToss))
    End Function
    Public Function isomorphicVariable(ByVal objVariable2 As qbVariable, _
                                       ByRef strExplanation As String) _
           As Boolean
        If Not checkUsable_("containedVariable", "Returning False") Then
            strExplanation = "Object is not usable"
            Return (False)
        End If
        With Me
            If Me.Dope.isUDT OrElse objVariable2.Dope.isUDT Then
                strExplanation = "Variables aren't isomorphic because " & _
                                 "either or both is a UDT"
                Return (False)
            End If
            If Not _OBJvariableType.isomorphicType(objVariable2.Dope, _
                                                   strExplanation) Then
                strExplanation = "Variables aren't isomorphic because " & _
                                 "their types aren't isomorphic"
                Return (False)
            End If
            strExplanation = "The types of the variables are isomorphic"
            Dim booStringIdentical As Boolean = .stringIdentical(objVariable2)
            strExplanation &= strExplanation & ": " & _
                              "the variables themselves are " & _
                              CStr(IIf(booStringIdentical, "", "not ")) & " " & _
                              "isomorphic: " & _
                              "their string representations " & _
                              CStr(IIf(booStringIdentical, "", "do not ")) & " " & _
                              "match"
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Test for a variable of scalar type, or, of Variant type and scalar
    ' contents
    '
    '
    Public Function isScalar() As Boolean
        If Not checkUsable_("isScalar", "Returning False") Then Return (False)
        With Me.Dope
            If .isScalar Then Return (True)
            If Not .isVariant Then Return (False)
            Return (isScalarType(.VarType))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Make a random array as a fromString expression
    '
    '
    Public Shared Function mkRandomVariable() As String
        Dim objArrayType As qbVariableType.qbVariableType = _
            New qbVariableType.qbVariableType(_OBJvariableType.mkRandomType)
        With objArrayType
            If .isUnknown OrElse .isNull Then Return .ToString
            Return (.ToString & _
                    ":" & _
                    mkRandomVariable_addValues_(objArrayType, 0))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Recursive procedure to populate the random array on behalf of
    ' mkRandomArray
    '
    '
    Private Shared Function mkRandomVariable_addValues_(ByVal objType As qbVariableType.qbVariableType, _
                                                        ByVal intRecursion As Integer) _
            As String
        If Rnd() <= 0.1 Then Return ("*")
        With objType
            If .isArray _
               OrElse _
               .isVariant AndAlso .VarType = ENUvarType.vtArray Then
                Return (mkRandomArrayValues_(objType, 0))
            ElseIf .isUDT _
                   OrElse _
                   .isVariant AndAlso .VarType = ENUvarType.vtUDT Then
                Return (mkRandomUDTvalues_(objType))
            Else
                Return (_OBJvariableType.mkRandomScalarValue.ToString)
            End If
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Return the scalar value on behalf of mkRandomVariable
    '
    '
    Private Shared Function mkRandomVariable_mkScalar_ _
                            (ByVal objType As qbVariableType.qbVariableType) _
            As String
        With objType
            Dim strValue As String
            If .VarType = ENUvarType.vtVariant Then
                strValue = .mkRandomVariantValue
            Else
                strValue = _OBJutilities.object2String(.mkRandomScalarValue(.VarType.ToString), _
                                                        True)
                If UCase(Mid(strValue, 1, 7)) = "SYSTEM." Then
                    strValue = Mid(strValue, 8)
                End If
            End If
            Return (strValue)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Make the object not usable
    '
    '
    Public Function mkUnusable() As Boolean
        USRstate.booUsable = False
    End Function

    ' -----------------------------------------------------------------------
    ' Manufacture a variable from a from string
    '
    '
    Public Shared Function mkVariable(ByVal strFromString As String) As qbVariable
        Dim objNew As qbVariable
        Try
            objNew = New qbVariable(strFromString)
        Catch
            errorHandler_("Cannot create new qbVariable using the from string " & _
                          _OBJutilities.enquote(strFromString) & ": " & _
                          Err.Number & " " & Err.Description, _
                          ClassName, "mkVariable", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        Return (objNew)
    End Function

    ' -----------------------------------------------------------------------
    ' Manufacture a variable from a from string
    '
    '
    Public Shared Function mkVariableFromValue(ByVal objValue As Object) As qbVariable
        Try
            Dim strValue As String = CStr(objValue)
            Return (mkVariable(_OBJutilities.object2String(strValue, True)))
        Catch objEx As Exception
            errorHandler_("Unable to make a variable from a value: " & _
                          Err.Number & " " & Err.Description, _
                          ClassName, _
                          "mkVariableFromValue", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Return and alter name
    '
    '
    Public Property Name() As String
        Get
            Return (USRstate.strName)
        End Get
        Set(ByVal strNewValue As String)
            If Not checkUsable_("Name set", "No change made") Then Return
            USRstate.strName = strNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' .Net decorated string to Quick Basic decorated string
    '
    '
    ' Converts the format [SYSTEM.]netType(value) to qbType(value) or a 
    ' null string (with no other error indication) when the Net value cannot 
    ' be converted.
    '
    '
    Private Function netDeco2qbDeco_(ByVal strNetDeco As String) As String
        ' --- Scan the .Net decorated string
        ' Create a scan object
        Dim objScanner As qbScanner.qbScanner
        Try
            objScanner = New qbScanner.qbScanner
            objScanner.scan(strNetDeco)
        Catch
            errorHandler_("Cannot create scanner: " & _
                          Err.Number & " " & Err.Description, _
                          "netDeco2qb_", _
                          "Marking object unusable and returning a null string")
            Me.mkUnusable() : Return ("")
        End Try
        ' Use the scan object to convert formats      
        With objScanner
            If Not .scan(strNetDeco) Then Return ("")
            If .TokenCount < 4 Then Return ("")
            Dim intIndex As Integer = 1
            Dim strQBdeco As String
            If .checkToken(intIndex, "SYSTEM") Then
                .checkTokenByTypeName(intIndex, "Period")
            End If
            If Not .checkTokenByTypeName(intIndex, "IDENTIFIER") Then Return ("")
            If Not .checkToken(intIndex, "(") Then Return ("")
            Dim intEndIndex As Integer = .findRightParenthesis(intIndex)
            If intEndIndex = 0 Then Return ("")
            Dim strQBtype As String
            Try
                strQBtype = netType2QBdomain(.sourceMid(intIndex - 2)).ToString
            Catch
                Return ("")
            End Try
            Return (strQBtype & "(" & .sourceMid(intIndex, intEndIndex - intIndex) & ")")
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Convert .Net value to a qbVariable
    '
    '
    Public Shared Function netValue2QBvariable(ByVal objNet As Object) As qbVariable
        Try
            Dim strNew As String = _OBJvariableType.netValue2QBdomain(objNet).ToString
            If _OBJvariableType.isScalarType(strNew) Then
                strNew &= ":" & CStr(objNet)
            End If
            Dim objQBV As qbVariable = New qbVariable(strNew)
            Return (objQBV)
        Catch
            errorHandler_("Cannot create qbVariable: " & Err.Number & " " & Err.Description, _
                          ClassName, _
                          "", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Convert state to XML
    '
    '
    Public Function object2XML() As String
        With USRstate
            Dim strValue As String
            If (TypeOf .objValue Is Collection) Then
                strValue = "(" & _
                           OBJcollectionUtilities.collection2String(CType(.objValue, Collection), _
                                                                    booReadable:=True, _
                                                                    booDeco:=.objDope.isVariant _
                                                                             OrElse _
                                                                             .objDope.VarType = ENUvarType.vtVariant) & _
                           ")"
            Else
                strValue = _OBJutilities.object2String(.objValue, True)
            End If
            Return (_OBJutilities.objectInfo2XML(Me.ClassName, _
                                                _OBJutilities.soft2HardParagraph(Me.About), _
                                                True, True, _
                                                "booUsable", _
                                                "Indicates the usability of the object", _
                                                CStr(USRstate.booUsable), _
                                                "strName", _
                                                "Identifies the object instance", _
                                                .strName, _
                                                "strVariableName", _
                                                "Identifies the variable", _
                                                .strVariableName, _
                                                "booVariableNameDefaults", _
                                                "True indicates that the variable name has default value", _
                                                CStr(.booVariableNameDefaults), _
                                                "objDope", _
                                                "Dope as a qbVariableType", _
                                                Replace(.objDope.object2XML, _
                                                        vbNewLine, _
                                                        vbNewLine & "    "), _
                                                "objValue", _
                                                "Value of this variable: " & _
                                                "parentheses indicate an array or UDT, " & _
                                                "represented as a collection", _
                                                strValue, _
                                                "objTag", _
                                                "User's tag", _
                                                _OBJutilities.object2String(.objTag)))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Reset the variable
    '
    '
    Public Function resetVariable() As Boolean
        If Not checkUsable_("resetVariable", "Variable cannot be reset") Then
            Return (False)
        End If
        With USRstate
            .objDope = _OBJvariableType.mkType(ENUvarType.vtUnknown)
            .objValue = Nothing
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Test for string identity
    '
    '
    Public Function stringIdentical(ByVal objValue2 As qbVariable) As Boolean
        If Not checkUsable_("stringIdentical", "Returning False") Then
            Return (False)
        End If
        Return (toStringData_(Me.ToString) = toStringData_(objValue2.ToString))
    End Function

    ' -----------------------------------------------------------------------
    ' Return and assign user's Tag
    '
    '
    Public Property Tag() As Object
        Get
            If Not checkUsable_("Tag get", "Returning Nothing") Then
                Return (Nothing)
            End If
            Return (USRstate.objTag)
        End Get
        Set(ByVal objNewValue As Object)
            If Not checkUsable_("Tag set", "No change made") Then
                Return
            End If
            USRstate.objTag = objNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Test the object
    '
    '
    ' --- Self test: makes a special test object
    Public Overloads Function test(ByRef strReport As String) As Boolean
        Return test(strReport, True)
    End Function
    ' --- Self test: allows user to specify option to create a test object
    Public Overloads Function test(ByRef strReport As String, _
                                   ByVal booMkObject As Boolean) As Boolean
        Dim datStart As Date = Now
        Try
            RaiseEvent msgEvent("Starting to test", 0)
            RaiseEvent msgEvent("Creating test object", 1)
            strReport = "Testing qbVariable at " & Now
            strReport &= vbNewLine & vbNewLine & "Creating test object"
            Dim objTest As qbVariable
            If booMkObject Then
                RaiseEvent msgEvent("Creating the test object", 1)
                Try
                    objTest = New qbVariable
                Catch
                    errorHandler_("Cannot create test object: " & _
                                Err.Number & " " & Err.Description, _
                                "test", _
                                "Marking main object unusable, returning False")
                    Me.mkUnusable()
                    Return False
                End Try
                RaiseEvent msgEvent("Test object has been created", 1)
            Else
                objTest = Me
            End If
            With objTest
                strReport = "Testing [" & objTest.Name & "] at " & Now
                ' --- Test the documentation examples
                Dim booInspection As Boolean = True
                Dim booOK As Boolean
                Dim strActivity As String = _
                    test__startActivity_("Testing documentation examples", _
                                         strReport)
                Dim intIndex1 As Integer
                Dim objTest2 As qbVariable = New qbVariable
                Dim strSplit() As String
                Try
                    strSplit = Split(TEST_EXAMPLES, vbNewLine & vbNewLine)
                Catch
                    errorHandler_("Cannot split: " & Err.Number & Err.Description, _
                                  ClassName, _
                                  "test_", _
                                  "Returning False")
                    Return (False)
                End Try
                Dim strInspection As String
                Dim strSplit2() As String
                Dim strTostring As String
                Dim strEnquoted As String
                For intIndex1 = 0 To UBound(strSplit)
                    Try
                        strSplit2 = Split(strSplit(intIndex1), vbNewLine)
                    Catch
                        errorHandler_("Cannot split: " & Err.Number & Err.Description, _
                                      ClassName, _
                                      "test_", _
                                      "Returning False")
                        Return (False)
                    End Try
                    strEnquoted = _OBJutilities.enquote(strSplit2(0))
                    RaiseEvent progressEvent(strActivity, _
                                                    "example", _
                                                    intIndex1 + 1, _
                                                    UBound(strSplit) + 1, _
                                                    2, _
                                                    "Example: " & strEnquoted)
                    strReport &= vbNewLine & vbNewLine & "Testing " & strEnquoted
                    .fromString(strSplit2(0))
                    strTostring = .ToString
                    strInspection = ""
                    booOK = .inspect(strInspection) AndAlso (strTostring = strSplit2(1))
                    If booOK Then
                        strInspection = ""
                    Else
                        strInspection = vbNewLine & vbNewLine & _
                                        _OBJutilities.string2Box _
                                        (.object2XML & _
                                         vbNewLine & vbNewLine & _
                                         strInspection)
                    End If
                    strReport &= vbNewLine & vbNewLine & _
                                    "fromString " & _
                                    _OBJutilities.enquote(strSplit2(0)) & " " & _
                                    CStr(IIf(booOK, "correctly results ", "fails to result ")) & " " & _
                                    "in toString " & _
                                    _OBJutilities.enquote(strSplit2(1)) & ": " & _
                                    CStr(IIf(booOK, _
                                             "", _
                                             "it results in toString " & _
                                             _OBJutilities.enquote(strTostring))) & _
                                    strInspection
                    If Not booOK Then
                        Return (False)
                    End If
                    .resetVariable()
                Next intIndex1
                RaiseEvent msgEvent("Documentation example test complete", -1)
                ' --- Test containment
                If TEST_CONTAINMENT <> "" Then
                    strActivity = test__startActivity_("Testing containment", strReport)
                    Try
                        strSplit = Split(TEST_CONTAINMENT, vbNewLine)
                    Catch
                        errorHandler_("Cannot split: " & Err.Number & Err.Description, _
                                    ClassName, _
                                    "test_", _
                                    "Returning False")
                        Return (False)
                    End Try
                    For intIndex1 = 0 To UBound(strSplit)
                        Try
                            strSplit2 = Split(strSplit(intIndex1), "\")
                        Catch
                            errorHandler_("Cannot split: " & Err.Number & Err.Description, _
                                        ClassName, _
                                        "test_", _
                                        "Returning False")
                            Return (False)
                        End Try
                        RaiseEvent progressEvent(strActivity, _
                                                        "test", _
                                                        intIndex1 + 1, _
                                                        UBound(strSplit) + 1, _
                                                        2, _
                                                        "")
                        .fromString(strSplit2(0))
                        objTest2.fromString(strSplit2(1))
                        Select Case UCase(strSplit2(2))
                            Case "CONTAINED"
                                booOK = .containedVariable(objTest2)
                            Case "ISOMORPHIC"
                                booOK = .isomorphicVariable(objTest2)
                            Case ""
                                booOK = Not .containedVariable(objTest2) _
                                        AndAlso _
                                        Not .containedVariable(objTest2)
                            Case Else
                                errorHandler_("Error in a test case", _
                                            ClassName, _
                                            "test_", _
                                            "Returning False")
                                Return (False)
                        End Select
                        booInspection = booInspection AndAlso booOK
                        strReport &= vbNewLine & vbNewLine & _
                                    "The expected relationship " & _
                                    _OBJutilities.enquote(strSplit2(2)) & _
                                    "does " & _
                                    CStr(IIf(booOK, "", "NOT ")) & _
                                    "obtain between " & _
                                    strSplit2(0) & " and " & strSplit2(1)
                    Next intIndex1
                End If
                If booMkObject Then
                    .dispose() : objTest = Nothing
                End If
            End With
        Catch
            strReport &= vbNewLine & vbNewLine & _
                         "Error occured in test: " & _
                         Err.Number & " " & Err.Description
            Return (False)
        End Try
        RaiseEvent msgEvent("The test has succeeded! It took about " & _
                            DateDiff("s", datStart, Now) & " second(s)", -1)
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Format the boxed XML and inspection report on behalf of test
    '
    '
    Private Shared Function test__formatReport_(ByVal strXML As String, _
                                                ByVal strInspection As String, _
                                                ByVal intSequence As Integer) _
            As String
        With _OBJutilities
            Return (.string2Box(.string2Box(strXML, _
                                           "Extended Markup Language (XML)") & _
                               vbNewLine & vbNewLine & _
                               .string2Box(strInspection, _
                                           "Inspection report"), _
                               "Q B A R R A Y   T E S T   " & intSequence))
        End With
    End Function

    ' ------------------------------------------------------------------------
    ' On a random basis and infrequently, remove the type and/or bounds to
    ' test fromString's ability to create sensible defaults
    '
    '
    Private Shared Function test__randomTrimmer_(ByVal strFromString As String) As String
        If Rnd() > 0.1 Then Return (strFromString)
        Dim dblRnd As Double = Rnd()
        If dblRnd <= 0.33 Then
            ' 33% of the time, take the type out
            Return (_OBJutilities.itemPhrase(strFromString, 1, -1, ",", False))
        ElseIf dblRnd <= 0.66 Then
            ' 33% of the time, take the bounds out
            Return (_OBJutilities.item(strFromString, 1, ",", False) & "," & _
                   _OBJutilities.item(strFromString, 2, ":", False))
        Else
            ' 33% of the time take both the type and the bounds out
            Return (_OBJutilities.item(strFromString, 2, ":", False))
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Start the testing activity
    '
    '
    Private Function test__startActivity_(ByVal strActivity As String, _
                                          ByRef strReport As String) _
            As String
        RaiseEvent msgEvent(strActivity, 1)
        strReport &= vbNewLine & vbNewLine & strActivity
        Return (strActivity)
    End Function

    ' -----------------------------------------------------------------------
    ' Return a description of the variable suitable for display
    '
    '
    Public Function toDescription() As String
        If Not checkUsable_("toDescription", _
                            "Returning error message") Then
            Return ("Object is not usable, no description is available")
        End If
        Return (USRstate.objDope.ToString & ": " & _
               CStr(IIf(Me.isClear, "is empty", "contains nondefault values")))
    End Function

    ' -----------------------------------------------------------------------
    ' Format two-dimensional array
    '
    '
    Public Overloads Function toMatrix() As String
        Return toMatrix(False)
    End Function
    Public Overloads Function toMatrix(ByVal booDeco As Boolean) As String
        If Not checkUsable_("toMatrix", "Returning a null string") Then Return ("")
        With USRstate
            If Not .objDope.isArray OrElse .objDope.Dimensions <> 2 Then
                errorHandler_("Variable is not a 2-dimension array", _
                              "", _
                              "No change made: null string returned")
                Return ("")
            End If
            Return (OBJcollectionUtilities.collection2Report(CType(.objValue, Collection), _
                                                             booDeco:=booDeco))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Convert to string
    '
    '
    Public Overrides Function toString() As String
        If Not checkUsable_("toString", "Returning a null string") Then
            Return ("")
        End If
        With USRstate
            Dim strValue As String = toString_(.objValue)
            If strValue <> "" Then strValue = ":" & strValue
            Return (.objDope.ToString & _
                   CStr(IIf(.objDope.isUnknown OrElse .objDope.isNull, _
                            "", _
                            strValue)))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Recursion on behalf of toString method
    '
    '
    Private Function toString_(ByVal objValue As Object) As String
        If (TypeOf objValue Is Collection) Then
            Dim strOutstring As String
            With CType(objValue, Collection)
                Dim booNondefault As Boolean
                Dim intIndex1 As Integer = 1
                Dim intIndex2 As Integer
                Dim strDefault As String = toString__scalarToString_(Me.Dope.defaultValue)
                Dim strNext As String
                For intIndex1 = 1 To .Count
                    If (TypeOf .Item(intIndex1) Is Collection) Then
                        strNext = toString_(.Item(intIndex1))
                        If strNext <> "*" Then booNondefault = True
                        strNext = "(" & strNext & ")"
                    ElseIf (TypeOf .Item(intIndex1) Is qbVariable) Then
                        strNext = CType(.Item(intIndex1), qbVariable).ToString
                        strNext = Mid(strNext, InStr(strNext & ":", ":") + 1)
                        If strNext <> strDefault Then booNondefault = True
                    Else
                        strNext = toString__scalarToString_(.Item(intIndex1))
                        If strNext <> strDefault Then booNondefault = True
                    End If
                    _OBJutilities.append(strOutstring, ",", strNext)
                Next intIndex1
                If booNondefault Then Return (strOutstring)
                Return ("*")
            End With
        Else
            Return (toString__scalarToString_(objValue))
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Converts the scalar to a string on behalf of toString method
    '
    '
    Private Function toString__scalarToString_(ByVal objValue As Object) As String
        Dim strScalar As String
        strScalar = _OBJutilities.object2String(objValue, True)
        Return (netDeco2qbDeco_(strScalar))
    End Function

    ' -----------------------------------------------------------------------
    ' Convert type info to string
    '
    '
    Public Function toStringTypeOnly() As String
        If Not checkUsable_("toStringtoStringTypeOnly", "Returning a null string") Then
            Return ("")
        End If
        Return (USRstate.objDope.ToString)
    End Function

    ' -----------------------------------------------------------------------
    ' Convert type info and value to string
    '
    '
    Public Function toStringWithType() As String
        If Not checkUsable_("toStringWithType", "Returning a null string") Then
            Return ("")
        End If
        Return (Me.toStringTypeOnly & "," & Me.ToString)
    End Function

    ' -----------------------------------------------------------------------
    ' Return the UDT member
    '
    '
    Public ReadOnly Property UDTmember(ByVal objMemberID As Object) As qbVariable
        Get
            If Not checkUsable_("UDTmember Get", "Returning Nothing") Then Return (Nothing)
            If Not Me.Dope.isUDT Then
                errorHandler_("The property is inapplicable unless the object is a UDT", _
                              "UDTmember Get", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            Dim intMemberIndex As Integer
            If (TypeOf objMemberID Is Integer) Then
                intMemberIndex = CInt(objMemberID)
            ElseIf (TypeOf objMemberID Is System.String) Then
                intMemberIndex = Me.Dope.udtMemberDeref(CStr(objMemberID))
            Else
                errorHandler_("Invalid index " & _OBJutilities.object2String(objMemberID) & " " & _
                              "is not an integer or a string", _
                              "UDTmember Get", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            If intMemberIndex < 1 OrElse intMemberIndex > Me.Dope.UDTmemberCount Then
                errorHandler_("Invalid index " & intMemberIndex & " " & _
                              "is zero, negative, or greater than member count " & _
                              Me.Dope.UDTmemberCount, _
                              "UDTmember Get", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            Return (CType(CType(USRstate.objValue, Collection).Item(intMemberIndex), _
                         qbVariable))
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Return the usability of the object
    '
    '
    Public ReadOnly Property Usable() As Boolean
        Get
            Return (USRstate.booUsable)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Obtain variable's value
    '
    '
    ' --- Obtains value of a scalar
    Public Overloads Function value() As Object
        Return (value(""))
    End Function
    ' --- Obtains value of a one-dimensional array
    Public Overloads Function value(ByVal intIndex As Integer) As Object
        Return (value(CStr(intIndex)))
    End Function
    ' --- Obtains value of a two-dimensional array
    Public Overloads Function value(ByVal intIndex1 As Integer, _
                                    ByVal intIndex2 As Integer) As Object
        Return (value(CStr(intIndex1) & "," & CStr(intIndex2)))
    End Function
    ' --- Obtains value of scalar or array
    Public Overloads Function value(ByVal strID As String) As Object
        If Not checkUsable_("value", "Returning Nothing") Then
            Return (qbVariableType.qbVariableType.ENUvarType.vtUnknown)
        End If
        With USRstate
            If Me.Dope.isUDT Then
                Return (derefMemberName(strID).value)
            End If
            If strID = "" AndAlso .objDope.Dimensions <> 0 _
               OrElse _
               strID <> "" _
               AndAlso _
               _OBJutilities.items(strID, ",", False) <> .objDope.Dimensions Then
                errorHandler_("Number of indexes specified does not correspond to dimensions", _
                              "value", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            If .objDope.isArray Then
                ' --- Access collection
                Return (OBJcollectionUtilities.dewey2Member(CType(.objValue, Collection), _
                                                           indexes2Dewey_(strID, .objDope)))
            ElseIf (TypeOf .objValue Is qbVariable) Then
                Dim objValue As qbVariable = CType(.objValue, qbVariable)
                If Not objValue.Dope.isScalar Then
                    errorHandler_("Variable value " & _
                                  _OBJutilities.object2String(.objValue) & " " & _
                                  "is not scalar", _
                                  "value", _
                                  "Making variable object unusable: returning Nothing")
                    Me.mkUnusable()
                    Return Nothing
                End If
                Return (objValue.value)
            Else
                Try
                    Dim objValue As Object = _OBJutilities.object2Scalar(.objValue)
                Catch
                    errorHandler_("Variable value " & _
                                  _OBJutilities.object2String(.objValue) & " " & _
                                  "is not scalar", _
                                  "value", _
                                  "Making variable object unusable: returning Nothing")
                    Me.mkUnusable()
                    Return Nothing
                End Try
                Return .objValue
            End If
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Modify value
    '
    '
    ' --- Modify scalar
    Public Overloads Function valueSet(ByVal objValue As Object) As Boolean
        Return (valueSet(objValue, "", ""))
    End Function
    ' --- Modify one-dimension array
    Public Overloads Function valueSet(ByVal objValue As Object, _
                                       ByVal intIndex As Integer) As Boolean
        Return (valueSet(objValue, "", CStr(intIndex)))
    End Function
    ' --- Modify two-dimension array
    Public Overloads Function valueSet(ByVal objValue As Object, _
                                       ByVal intIndex1 As Integer, _
                                       ByVal intIndex2 As Integer) As Boolean
        Return (valueSet(objValue, "", CStr(intIndex1) & CStr(intIndex2)))
    End Function
    ' --- Modify UDT member or array (which is not in a udt)
    Public Overloads Function valueSet(ByVal objValue As Object, _
                                       ByVal strIndexesOrMemberName As String) As Boolean
        If Not checkUsable_("valueSet", "Returning False") Then
            Return (False)
        End If
        With Me.Dope
            If .isUDT Then
                Return (valueSet(objValue, strIndexesOrMemberName, ""))
            ElseIf .isArray Then
                Return (valueSet(objValue, "", strIndexesOrMemberName))
            Else
                errorHandler_("This syntax cannot be used when the object does not " & _
                              "represent a UDT or an array", _
                              "valueSet", _
                              "Returning False")
            End If
        End With
    End Function
    ' --- Modify n-dimensional array which may be a UDT member
    Public Overloads Function valueSet(ByVal objValue As Object, _
                                       ByVal strMemberName As String, _
                                       ByVal strIndexes As String) As Boolean
        If Not checkUsable_("valueSet", "Returning False") Then
            Return (False)
        End If
        With USRstate
            If Me.Dope.isUDT Then
                Dim objMember As qbVariable = derefMemberName(strMemberName)
                If (objMember Is Nothing) Then Return (Nothing)
                Return (derefMemberName(strMemberName).valueSet(objValue, strIndexes))
            End If
            If .objDope.isArray Then
                Dim objValueSet As Object
                If .objDope.VarType = ENUvarType.vtVariant Then
                    objValueSet = objValue
                ElseIf .objDope.isScalar Then
                    objValueSet = _OBJutilities.canonicalTypeCast(objValue, _
                                                                  .objDope.qbDomain2NetType(.objDope.VarType))
                Else
                    errorHandler_("Cannot set array to vartype " & _
                                  .objDope.VarType.ToString, _
                                  "valueSet", _
                                  "No change made")
                    Return False
                End If
                Return _
                Not (OBJcollectionUtilities.collectionItemSet _
                    (CType(.objValue, Collection), _
                     indexes2Dewey_(strIndexes, .objDope), _
                     objValueSet) _
                     Is _
                     Nothing)
            ElseIf Trim(strMemberName) <> "" Then
                errorHandler_("Scalar cannot be indexed", "valueSet", "Scalar has not been changed")
                Return (False)
            Else
                Return (valueSet_(objValue, .objDope, .objValue))
            End If
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Return and assign the variable name
    '
    '
    Public Property VariableName() As String
        Get
            If Not checkUsable_("Name Get", "Returning a null string") Then
                Return ("")
            End If
            Return (USRstate.strVariableName)
        End Get
        Set(ByVal strNewValue As String)
            If Not checkUsable_("Name Set", "No change made") Then
                Return
            End If
            With USRstate
                .booVariableNameDefaults = False
                .strVariableName = strNewValue
            End With
            Return
        End Set
    End Property

    ' ***** Private procedures **********************************************

    ' -----------------------------------------------------------------------
    ' Check usability
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String, _
                                  ByVal strHelp As String) As Boolean
        If Not USRstate.booUsable Then
            errorHandler_("Object is not usable", _
                          strProcedure, _
                          strHelp)
            Return (False)
        End If
        Return (True)
    End Function

    ' ------------------------------------------------------------------------
    ' Clear the array values
    '
    '
    Private Function clearArrayCollectionValues_(ByVal colCollection As Collection, _
                                                 ByVal objValue As Object, _
                                                 ByVal intRecursion As Integer) As Boolean
        Dim intIndex1 As Integer
        With colCollection
            For intIndex1 = 1 To .Count
                RaiseEvent progressEvent("Clearing the array's collection", _
                                            "item", _
                                            intIndex1, _
                                            .Count, _
                                            intRecursion, _
                                            "")
                If (TypeOf .Item(intIndex1) Is Collection) Then
                    If Not clearArrayCollectionValues_(CType(.Item(intIndex1), Collection), _
                                                       objValue, _
                                                       intRecursion + 1) Then
                        Exit For
                    End If
                Else
                    Try
                        .Remove(intIndex1)
                        .Add(objValue, , intIndex1)
                    Catch
                        errorHandler_("Could not clear array collection entry", _
                                      "clearArrayCollectionValues_", _
                                      "Not changing rest of collection: returning False")
                        Exit For
                    End Try
                End If
            Next intIndex1
            Return (intIndex1 > .Count)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Clear UDT collection
    '
    '
    Private Function clearUDTcollectionValues_(ByRef colCollection As Collection) _
            As Boolean
        With colCollection
            Dim intIndex1 As Integer
            Dim objNext As qbVariable
            For intIndex1 = .Count To 1 Step -1
                RaiseEvent progressEvent("Clearing the UDT's collection", _
                                         "item", _
                                         intIndex1, _
                                         .Count, _
                                         0, _
                                         "")
                objNext = CType(.Item(intIndex1), qbVariable)
                With objNext
                    .dispose()
                    objNext = Nothing
                End With
                .Remove(intIndex1)
            Next intIndex1
            If intIndex1 > 0 Then Return (False)
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Return depth of object considered as an orthogonal collection
    '
    '
    Private Function collectionDepth_(ByVal objObject As Object) As Integer
        Dim objWork As Object = objObject
        Dim intDepth As Integer
        Do
            If Not (TypeOf objWork Is Collection) Then Return intDepth
            intDepth += 1
            Try
                objWork = CType(objWork, Collection).Item(1)
            Catch
                objWork = Nothing
            End Try
        Loop
    End Function

    ' -----------------------------------------------------------------------
    ' Return the "empirical" dope
    '
    '
    ' This method determines, by examination, the structure of the 
    ' array that is represented in objValue. It returns this structure as a 
    ' qbVariableType.
    '
    ' If the objValue does not represent an array this method returns Nothing.
    '
    '
    Private Function empiricalDope_(ByVal objValue As Object) As qbVariableType.qbVariableType
        Dim objDope As qbVariableType.qbVariableType
        Try
            objDope = New qbVariableType.qbVariableType
        Catch
            errorHandler_("Cannot create dope object", _
                          "empiricalDope_", _
                          "Returning Nothing: making object unusable")
            Me.mkUnusable()
            Return (Nothing)
        End Try
        With objDope
            .fromString("Unknown")
            If (objValue Is Nothing) Then Return (objDope)
            If (TypeOf objValue Is Collection) Then
                Dim colHandle As Collection = CType(objValue, Collection)
                Dim strUDTtypes As String
                If inspectUDTCollection_(colHandle, strUDTtypes) Then
                    ' --- UDT dope
                    .fromString("UDT," & strUDTtypes)
                Else
                    ' --- Array dope
                    Dim colType As Collection
                    colType = OBJcollectionUtilities.collectionTypes(colHandle, booDrillDown:=True)
                    If (colType Is Nothing) Then
                        errorHandler_("Cannot determine entry type", _
                                    "empiricalDope_", _
                                    "Returning Nothing: making object unusable")
                        Me.mkUnusable()
                        objDope.dispose()
                        Return (Nothing)
                    End If
                    Dim strType As String = "Variant"
                    With colType
                        If .Count = 1 Then strType = _
                            qbVariableType.qbVariableType.netType2QBdomain(CStr(colType.Item(1))).ToString
                    End With
                    .fromString("Array," & strType & "," & _
                                empiricalDope__bounds_(colHandle))
                End If
            Else
                .fromString(_OBJvariableType.netType2QBdomain(objValue.GetType.ToString).ToString)
            End If
            Return (objDope)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Empirically determine array bounds
    '
    '
    Private Function empiricalDope__bounds_(ByVal objValue As Collection) As String
        With objValue
            Dim strBounds As String
            _OBJutilities.append(strBounds, ",", "0," & CStr(.Count - 1))
            If (TypeOf .Item(1) Is Collection) Then
                _OBJutilities.append(strBounds, _
                                     ",", _
                                     empiricalDope__bounds_(CType(.Item(1), Collection)))
            End If
            Return (strBounds)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Interface to error handler
    '
    '
    ' --- Stateful
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, USRstate.strName, strProcedure, strHelp)
    End Sub
    ' --- Shared
    Private Overloads Shared Sub errorHandler_(ByVal strMessage As String, _
                                               ByVal strObject As String, _
                                               ByVal strProcedure As String, _
                                               ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, strObject, strProcedure, strHelp)
    End Sub

    ' -----------------------------------------------------------------------
    ' Locate comma at same parenthesis level
    '
    '
    Private Function findBalancedComma_(ByVal objScanner As qbScanner.qbScanner, _
                                        ByVal intIndex As Integer, _
                                        ByVal intEndIndex As Integer) As Integer
        Dim intIndex1 As Integer = intIndex
        Do While intIndex1 <= intEndIndex
            Dim strNext As String = objScanner.sourceMid(intIndex1, 1)
            If strNext = "(" Then
                intIndex1 = objScanner.findRightParenthesis(intIndex1 + 1)
            ElseIf strNext = "," Then
                Exit Do
            End If
            intIndex1 += 1
        Loop
        Return intIndex1
    End Function

    ' -----------------------------------------------------------------------
    ' Calculates column index
    '
    '
    ' Where the index list is d(D),d(D-1)...d(1), this method returns the
    ' corresponding collection Dewey number, consisting of the "d" values
    ' separated by periods and adjusted for lowerBound.
    '
    '
    Private Function indexes2Dewey_(ByVal strIndexes As String, _
                                    ByVal objDope As qbVariableType.qbVariableType) _
            As String
        Dim strSplit() As String
        Try
            strSplit = Split(strIndexes, ",")
        Catch
            errorHandler_("Error in split: " & Err.Number & " " & Err.Description, _
                          "", _
                          "Returning 1: marking object unusable")
            Me.mkUnusable()
            Return ("1")
        End Try
        Dim intD As Integer
        Dim intIndex1 As Integer
        Dim strDewey As String
        With objDope
            For intIndex1 = 0 To UBound(strSplit)
                Try
                    intD = CInt(strSplit(intIndex1))
                Catch
                    errorHandler_("Index " & _OBJutilities.enquote(strSplit(intIndex1)) & " is not valid", _
                                  "indexes2colIndex_", _
                                  "Returning a collection index of 1")
                    Exit For
                End Try
                _OBJutilities.append(strDewey, _
                                     ".", _
                                     CStr(intD - objDope.LowerBound(intIndex1 + 1) + 1))
            Next intIndex1
            Return (strDewey)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Internal inspection
    '
    '
    Private Function inspection_() As Boolean
        Dim strReport As String
        If Not Me.inspect(strReport) Then
            errorHandler_("Internal inspection failed" & vbNewLine & vbNewLine & _
                          strReport, _
                          "inspection_", _
                          "Object is not usable")
            Return (False)
        End If
        Return (True)
    End Function

    ' ------------------------------------------------------------------------
    ' Inspect the UDT collection
    '
    '
    Private Overloads Function inspectUDTCollection_(ByVal colUDT As Collection) _
            As Boolean
        Dim strUDTtypes As String
        Return (inspectUDTCollection_(colUDT, strUDTtypes))
    End Function
    Private Overloads Function inspectUDTCollection_(ByVal colUDT As Collection, _
                                                     ByRef strUDTtypes As String) _
            As Boolean
        With colUDT
            Dim intIndex1 As Integer
            strUDTtypes = ""
            For intIndex1 = 1 To .Count
                If Not (TypeOf .Item(intIndex1) Is qbVariable) Then Exit For
                _OBJutilities.append(strUDTtypes, _
                                     ",", _
                                     "(" & CType(.Item(intIndex1), qbVariable).ToString & ")")
            Next intIndex1
            Return (intIndex1 > .Count)
        End With
    End Function

    ' ------------------------------------------------------------------------
    ' Tell caller if qbVariableTypes are "isomorphic dopes" (sigh)
    '
    '
    Private Function isomorphicDope_(ByVal objType1 As qbVariableType.qbVariableType, _
                                     ByVal objType2 As qbVariableType.qbVariableType) _
            As Boolean
        Return (_OBJvariableType.containedType(objType1, objType2) _
               AndAlso _
               _OBJvariableType.containedType(objType2, objType1))
    End Function

    ' -----------------------------------------------------------------
    ' Return the member type
    '
    '
    Private Function memberType_(ByVal objExpectedType As qbVariableType.qbVariableType, _
                                 ByVal intElementIndex As Integer) _
            As qbVariableType.qbVariableType
        Dim objMemberType As qbVariableType.qbVariableType
        If objExpectedType.isArray Then
            objMemberType = objExpectedType.arraySlice
        ElseIf objExpectedType.isUDT Then
            objMemberType = objExpectedType.udtMemberType(intElementIndex)
        Else
            objMemberType = objExpectedType.clone
        End If
        Return (objMemberType)
    End Function

    ' ------------------------------------------------------------------------
    ' Make the orthogonal array collection
    '
    '
    Private Function mkArrayCollection_(ByVal objDope As qbVariableType.qbVariableType, _
                                        ByVal intRecursion As Integer) _
            As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch
            errorHandler_("Cannot create array collection: " & Err.Number & " " & Err.Description, _
                          "mkArrayCollection", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        With objDope
            Dim intIndex1 As Integer
            Dim objInnerDope As qbVariableType.qbVariableType
            Dim objNew As Object
            Dim strInnerDope As String
            For intIndex1 = 1 To .BoundSize(1)
                RaiseEvent progressEvent("Making the array structure as an orthogonal collection", _
                                         "item", _
                                         intIndex1, _
                                         .BoundSize(1), _
                                         intRecursion, _
                                         "")
                If .Dimensions = 1 Then
                    objNew = .defaultValue
                Else
                    Try
                        objInnerDope = .clone
                        strInnerDope = .ToString
                        objInnerDope.fromString(_OBJutilities.itemPhrase(strInnerDope, _
                                                                         1, 2, ",", False) & "," & _
                                                _OBJutilities.itemPhrase(strInnerDope, _
                                                                         5, -1, ",", False))
                        objNew = mkArrayCollection_(objInnerDope, intRecursion + 1)
                    Catch
                        errorHandler_("Cannot create inner array collection: " & _
                                      Err.Number & " " & Err.Description, _
                                      "mkArrayCollection", _
                                      "Returning Nothing")
                        Try
                            objInnerDope.dispose()
                        Catch : End Try
                        Exit Function
                    End Try
                End If
                Try
                    colNew.Add(objNew)
                Catch
                    errorHandler_("Cannot extend array collection: " & _
                                  Err.Number & " " & Err.Description, _
                                  "mkArrayCollection", _
                                  "Returning Nothing")
                    Exit Function
                End Try
            Next intIndex1
            If intIndex1 <= .BoundSize(1) Then Return (Nothing)
            Return (colNew)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Return random values for an array
    '
    '
    Private Shared Function mkRandomArrayValues_(ByVal objArrayType As qbVariableType.qbVariableType, _
                                                 ByVal intRecursion As Integer) _
            As String
        With objArrayType
            Dim intIndex1 As Integer
            Dim strNext As String
            Dim strOutString As String = ""
            Dim objSlice As qbVariableType.qbVariableType = objArrayType.arraySlice
            Dim strRepeat As String
            For intIndex1 = .LowerBound(1) To .UpperBound(1)
                If .Dimensions = 1 Then
                    strNext = mkRandomVariable_mkScalar_(objArrayType)
                Else
                    strNext = "(" & _
                              mkRandomArrayValues_(objSlice, intRecursion + 1) & _
                              ")"
                End If
                strOutString = _OBJutilities.append(strOutString, ",", strNext)
                strRepeat = ""
                If Rnd() <= 0.4 Then
                    strRepeat = CStr(CInt(Rnd() * .BoundSize(.Dimensions)))
                ElseIf Rnd() <= 0.05 Then
                    strRepeat = "*"
                End If
                If strRepeat <> "" Then
                    strOutString = _OBJutilities.append(strOutString, _
                                                        ",", _
                                                        "(" & strRepeat & ")")
                End If
                RaiseEvent progressEventShared("Building the fromString expression for a random array", _
                                                CStr(IIf(.Dimensions = 1, "element", "slice")), _
                                                intIndex1 - .LowerBound(1) + 1, _
                                                .UpperBound(1) - .LowerBound(1) + 1, _
                                                intRecursion + 1, _
                                                "at dimension " & intRecursion + 1 & ": " & _
                                                "at index " & intIndex1 & " " & _
                                                "in an array with " & _
                                                .Dimensions & " " & _
                                                "dimension(s) and with bounds " & _
                                                .LowerBound(1) & " to " & .UpperBound(1) & " " & _
                                                "at dimension 1: " & _
                                                "fromString is now " & _
                                                _OBJutilities.enquote(_OBJutilities.ellipsis(strOutString, 64)))
            Next intIndex1
            Return (strOutString)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Return random values for an array
    '
    '
    Private Shared Function mkRandomUDTvalues_(ByVal objUDTType As qbVariableType.qbVariableType) _
            As String
        With objUDTType
            Dim intIndex1 As Integer
            Dim strNext As String
            Dim strOutString As String = ""
            For intIndex1 = 1 To .UDTmemberCount
                With .innerType(intIndex1)
                    If .isScalar Then
                        strNext = _OBJvariableType.mkRandomScalar
                    ElseIf .isArray Then
                        strNext = mkRandomArrayValues_(.innerType, 0)
                    ElseIf .isUDT Then
                        strNext = mkRandomUDTvalues_(.innerType)
                    ElseIf .isVariant Then
                        strNext = _OBJvariableType.mkRandomVariantValue
                    End If
                End With
                strOutString = _OBJutilities.append(strOutString, ",", strNext)
                RaiseEvent progressEventShared("Building the fromString expression for a random UDT", _
                                                "member", _
                                                intIndex1, _
                                                .UDTmemberCount, _
                                                0, _
                                                "")
            Next intIndex1
            Return (strOutString)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Create the collection for the UDT
    '
    '
    Private Function mkUDTcollection_(ByVal objDope As qbVariableType.qbVariableType) _
            As Collection
        Dim colUDT As Collection
        Try
            colUDT = New Collection
        Catch
            errorHandler_("Cannot create the UDT collection: " & _
                          Err.Number & " " & Err.Description, _
                          "mkUDTcollection_", _
                          "Making object unusable: returning Nothing")
            Me.mkUnusable() : Return (Nothing)
        End Try
        With objDope
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .UDTmemberCount
                Try
                    colUDT.Add(New qbVariable(.innerType(intIndex1).ToString))
                Catch ex As Exception
                    errorHandler_("Cannot extend the UDT collection: " & _
                                  Err.Number & " " & Err.Description, _
                                  "mkUDTcollection_", _
                                  "Making object unusable: returning Nothing")
                    Me.mkUnusable()
                    OBJcollectionUtilities.collectionClear(colUDT)
                    Return (Nothing)
                End Try
            Next intIndex1
            Return (colUDT)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Create a default variable name in the format <hungarianPrefix>nnnn
    ' or 
    '
    '
    Private Function mkVariableName_(ByVal objDope As qbVariableType.qbVariableType) _
            As String
        With objDope
            Dim strName As String
            strName = _OBJvariableType.hungarianPrefix(.VariableType)
            If .isArray OrElse .isVariant Then
                strName &= _OBJutilities.properCase(_OBJvariableType.vtPrefixRemove(.VarType.ToString))
            End If
            Return (strName & sequenceNumber_())
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Create the variable as a .Net scalar or Collection with default value
    '
    '
    Private Function mkVariableValue_(ByVal objDope As qbVariableType.qbVariableType) _
            As Object
        With objDope
            If .isArray OrElse .isVariant AndAlso .VarType = ENUvarType.vtArray Then
                ' --- Make the collection for the array
                Return (mkArrayCollection_(objDope, 0))
            ElseIf .isUDT OrElse .isVariant AndAlso .VarType = ENUvarType.vtUDT Then
                Return (mkUDTcollection_(objDope))
            Else
                ' --- Assign the .Net scalar default
                Return (.defaultValue)
            End If
        End With
    End Function

    ' --------------------------------------------------------------------
    ' Interfaces utilities.object2Scalar: handles Boolean values
    '
    '
    Private Shared Function object2Scalar_(ByVal objValue As Object) As Object
        If (TypeOf objValue Is Boolean) Then Return objValue
        Return (_OBJutilities.object2Scalar(objValue))
    End Function

    ' --------------------------------------------------------------------
    ' Format four-digit sequence number for Name and for VariableName
    '
    '
    Private Shared Function sequenceNumber_() As String
        Return (_OBJutilities.alignRight(CStr(Interlocked.Increment(_INTsequence)), 4, "0"))
    End Function

    ' --------------------------------------------------------------------
    ' Return the data part of the toString info
    '
    '
    Private Function toStringData_(ByVal strToString As String) As String
        Return (_OBJutilities.itemPhrase(strToString, 2, -1, ":", False))
    End Function

    ' -----------------------------------------------------------------
    ' Set Quick Basic value
    '
    '
    ' --- Call supplies the expected type
    Private Overloads Function valueSet_(ByVal objNetValue As Object, _
                                         ByVal objExpectedType As qbVariableType.qbVariableType, _
                                         ByRef objQBvalue As Object) As Boolean
        Return (valueSet__(objNetValue, objExpectedType, objQBvalue))
    End Function
    ' --- Call supplies only the expected type as an enumerator
    Private Overloads Function valueSet_(ByVal objNetValue As Object, _
                                         ByVal objExpectedType As _
                                               qbVariableType.qbVariableType.ENUvarType, _
                                         ByRef objQBvalue As Object) As Boolean
        Return (valueSet__(objNetValue, objExpectedType, objQBvalue))
    End Function
    ' --- Common logic
    Private Overloads Function valueSet__(ByVal objNetValue As Object, _
                                          ByVal objExpectedType As Object, _
                                          ByRef objQBvalue As Object) As Boolean
        Dim enuExpectedType As qbVariableType.qbVariableType.ENUvarType
        Dim objDope As qbVariableType.qbVariableType
        If (TypeOf objExpectedType Is qbVariableType.qbVariableType) Then
            objDope = CType(objExpectedType, qbVariableType.qbVariableType)
            enuExpectedType = objDope.VariableType
        Else
            enuExpectedType = CType(objExpectedType, qbVariableType.qbVariableType.ENUvarType)
        End If
        Try
            Select Case enuExpectedType
                Case ENUvarType.vtBoolean
                    objQBvalue = Nothing
                    If (TypeOf objNetValue Is String) Then
                        Select Case UCase(CStr(objNetValue))
                            Case "TRUE" : objQBvalue = True
                            Case "FALSE" : objQBvalue = False
                            Case Else
                                errorHandler_("Cannot assign the string .Net value " & _
                                                _OBJutilities.object2String(objNetValue) & " " & _
                                                "to the Boolean variable with the VariableName " & _
                                                Me.VariableName & " " & _
                                                "because the .Net value is not True or False", _
                                                "valueSet", _
                                                "Returning false")
                                Return (False)
                        End Select
                    End If
                    If (objQBvalue Is Nothing) Then objQBvalue = CBool(CDbl(objNetValue) <> 0)
                Case ENUvarType.vtByte
                    objQBvalue = CByte(objNetValue)
                Case ENUvarType.vtDouble
                    objQBvalue = CDbl(objNetValue)
                Case ENUvarType.vtInteger
                    objQBvalue = CShort(objNetValue)
                Case ENUvarType.vtLong
                    objQBvalue = CInt(objNetValue)
                Case ENUvarType.vtNull
                    If Not (objDope Is Nothing) Then
                        objDope.dispose()
                        objDope = _OBJvariableType.mkType(_OBJvariableType.netValue2QBdomain(objNetValue))
                    End If
                    Me.valueSet(objNetValue)
                Case ENUvarType.vtSingle
                    objQBvalue = CSng(objNetValue)
                Case ENUvarType.vtString
                    Dim strValue As String = CStr(objNetValue)
                    objQBvalue = strValue
                Case ENUvarType.vtUnknown
                    Dim objOldDope As qbVariableType.qbVariableType = objDope
                    objDope = _OBJvariableType.mkType(_OBJvariableType.netValue2QBdomain(objNetValue))
                    With Me
                        .Dope = objDope
                        .valueSet(objNetValue)
                        objQBvalue = .value
                    End With
                    objOldDope.dispose()
                Case ENUvarType.vtVariant
                    Dim strType As String = _
                        _OBJvariableType.netType2QBdomain(objNetValue.GetType.ToString).ToString
                    Me.fromString("Variant," & strType)
                    objQBvalue = _OBJvariableType.netValue2QBvalue(objNetValue, strType)
                Case Else
                    errorHandler_("Unexpected variable type", _
                                    "valueSet", _
                                    "Scalar has not been changed")
                    Return (False)
            End Select
        Catch
            errorHandler_("Cannot assign the .Net value " & _
                            _OBJutilities.object2String(objNetValue) & " " & _
                            "to the variable with the VariableName " & _
                            Me.VariableName, _
                            "valueSet", _
                            "Returning false")
            Return (False)
        End Try
        Return (True)
    End Function

End Class

