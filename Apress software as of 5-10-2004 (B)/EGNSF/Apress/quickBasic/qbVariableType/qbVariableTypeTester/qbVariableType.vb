Option Strict On

Imports System.Threading

Public Class qbVariableType

    Implements IComparable
    Implements ICloneable
    Implements IDisposable
    
    ' ***********************************************************************
    ' *                                                                     *
    ' * quickBasicEngine Variable Type                                      *
    ' *                                                                     *
    ' *                                                                     *
    ' * This class represents the type of a quickBasicEngine variable,      *
    ' * including support for an unknown type and Shared methods for        *
    ' * relating .Net types to Quick Basic types.                           * 
    ' *                                                                     *
    ' * The rest of this block describes:                                   *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  The variable types of the quickBasicEngine                  *
    ' *      *  Containment and isomorphism of variable types               *
    ' *      *  Properties, methods and events of this class                *
    ' *      *  Use as a shared tool source                                 *
    ' *      *  Compile-time symbols                                        *
    ' *      *  Type specification for fromString/from toString: BNF        *
    ' *      *  Cache considerations                                        * 
    ' *      *  Multithreading considerations                               * 
    ' *                                                                     *
    ' *                                                                     *
    ' * THE VARIABLE TYPES OF THE QUICKBASICENGINE ------------------------ *
    ' *                                                                     *
    ' * The variable types of the quickBasicEngine as supported by this     *
    ' * object fall into these general classes:                             *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  SCALARS are ordinary values with no structure as such, and  *
    ' *         they can have the type Boolean, Byte, Integer, Long, Single,*
    ' *         or String.                                                  *
    ' *                                                                     *
    ' *      *  VARIANTS are variables capable of "containing" variables    *
    ' *         including any scalar and even an array. In Quick Basic and  *
    ' *         this implementation, variants cannot contain variants. The  *
    ' *         elegance of such an idea is totally outweighed by its       *
    ' *         uselessness.                                                *
    ' *                                                                     *
    ' *      *  ARRAYS are variables with 1..n dimensions. At each dimension*
    ' *         an array has in Quick Basic and this implementation         *
    ' *         completely flexible lower and upper bounds. Flexible lower  *
    ' *         bounds are another nearly useless idea but here they were a *
    ' *         part of both Quick Basic and Visual Basic up to .Net. We    *
    ' *         must implement them ("Ess muss sein? Muss ess sein." -      *
    ' *         Beethoven.)                                                 *                                                                                  
    ' *                                                                     *
    ' *      *  UDTs (User Data Types) are variables that contain 1..n      *
    ' *         members, which may be a mix of scalars, variants, or arrays,*
    ' *         but cannot be nested UDTs.                                  *
    ' *                                                                     *
    ' *      *  The UNKNOWN variable type is what its name implies, a       *
    ' *         type we don't know. In this implementation the Unknown type *
    ' *         is assigned to the variable type in the constructor. In a   *
    ' *         planned future EGN implementation of a language for         *
    ' *         symbolic computation (which will probably be called FOG)    *
    ' *         this will be used to actually calculate with mystery values.*                                                              
    ' *                                                                     *
    ' *      *  The NULL variable type is primarily for assignment as the   *
    ' *         initial value of a variant value.                           *
    ' *                                                                     *
    ' *                                                                     *
    ' * The following types are exposed by this object.                     *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  ENUvarType.vtBoolean: True or False: default is False.      *
    ' *                                                                     *
    ' *      *  ENUvarType.vtByte: unsigned integer in the range 0..255:    *
    ' *         default is 0.                                               *
    ' *                                                                     *
    ' *      *  ENUvarType.vtInteger: integer in the range -32768..32767:   *
    ' *         note that this is different from VB.Net and like VB-6, and  *
    ' *         that our Integer is a Short in VB.Net: default is 0.        *
    ' *                                                                     *
    ' *      *  ENUvarType.vtLong: integer in the range -2^31..2^31-1:      *
    ' *         note that this is different from VB.Net and like VB-6, and  *
    ' *         that our Long is an Integer in VB.Net: default is 0.        *
    ' *                                                                     *
    ' *      *  ENUvarType.vtSingle: real number in the Single precision    *
    ' *         range of VB.Net.  Note that this is not fully compatible    *
    ' *         with Microsoft's old Quick Basic. The default value of      *
    ' *         Single is 0.                                                *
    ' *                                                                     *
    ' *      *  ENUvarType.vtDouble: real number in the Double precision    *
    ' *         range of VB.Net.  Note that this is not fully compatible    *
    ' *         with Microsoft's old Quick Basic. The default value of      *
    ' *         Double is 0.                                                *
    ' *                                                                     *
    ' *      *  ENUvarType.vtString: string restricted to 64K bytes when the*
    ' *         quickBasicEngine is compiled with QUICKBASICENGINE_EXTENSION*
    ' *         set to False: string restricted to the VB.Net string limit  *
    ' *         when this compile-time symbol is True. The default value of *
    ' *         string is the null string.                                  * 
    ' *                                                                     *
    ' *      *  ENUvarType.vtVariant: variant proto-object container for    *
    ' *         another value, which can be any type (including Array)      *
    ' *         except Variant itself. The default value of Variant is      *
    ' *         Nothing.                                                    *    
    ' *                                                                     *
    ' *         Most variants will occur set to a contained type; but the   *
    ' *         special "abstract" variant exists as a valid special state  *
    ' *         of this data type for variants inside arrays. This is a     *
    ' *         vtVariant variable type for which the Abstract property will*
    ' *         return True.                                                *
    ' *                                                                     *
    ' *      *  ENUvarType.vtArray: array of any dimensionality: no default.*
    ' *                                                                     *
    ' *         Unlike the variant, we do not support an abstract array.    *
    ' *         This is because an array in this implementation of Quick    *
    ' *         Basic is always an array "of" a definite type specified for *
    ' *         each entry, including a variant type...which is abstract.   *
    ' *                                                                     *
    ' *      *  ENUvarType.vtUDT: a user data type container for 1..n       *
    ' *         scalars, variants and/or arrays.                            *
    ' *                                                                     *
    ' *      *  ENUvarType.vtUnknown: default special value: no default for *
    ' *         this "default" is defined.                                  * 
    ' *                                                                     *
    ' *      *  ENUvarType.vtNull: used primarily for certain Variants: a   *
    ' *         dummy, uninitialized value: no default.                     *
    ' *                                                                     *
    ' *                                                                     *
    ' * When an Array, Variant, or UDT is represented this class also       *
    ' * contains the variant's type, the array type and array dimensions, or*
    ' * the collection of member types. These are delegates within the main *
    ' * object.                                                             *
    ' *                                                                     *
    ' *                                                                     *
    ' * CONTAINMENT AND ISOMORPHISM OF VARIABLE TYPES --------------------- *
    ' *                                                                     *
    ' * The containedType and the isomorphicType methods of this class      *
    ' * ensure that one type can be converted safely to another type.       *
    ' *                                                                     *
    ' * Type a is contained in type b when:                                 *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  Type a and type b are scalars (Boolean, Byte, Integer,      *
    ' *         Long, Single, Double or String) and all possible values     *
    ' *         of type a convert without error to type b                   *
    ' *                                                                     *
    ' *      *  Type a and type b are arrays. Each dimension, of type b,    *
    ' *         contains the same number of elements as the corresponding   *
    ' *         dimension of a, or more elements. The array entry type      *
    ' *         of a is contained in the array entry type of b according    *
    ' *         to this overall definition.                                 *
    ' *                                                                     *
    ' *         Note that lowerBounds of a and b may differ.                *
    ' *                                                                     *
    ' *      *  a is a scalar, a variant, or an array, and b is a Variant   *
    ' *                                                                     *
    ' *                                                                     *
    ' * Type a and type b are isomorphic types when a is contained in b,    *
    ' * and b is contained in a. Note that scalar types are isomorphic only *
    ' * when identical but array definitions may differ in lower bounds.    *
    ' *                                                                     *
    ' * If either a or b is a User Data Type, Null or Unknown, the types are*
    ' * never considered to contain each other.                             *
    ' *                                                                     *
    ' *                                                                     *
    ' * THE PROPERTIES, METHODS, AND EVENTS OF THIS CLASS ----------------- *
    ' *                                                                     *
    ' * Properties are named starting with an upper case letter; methods    *
    ' * and events use camelCase and start with a lower case letter; events *
    ' * end in the word Event.                                              *
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
    ' * interfaces: see the clone, compareTo and dispose methods, below.    *
    ' *                                                                     *
    ' *                                                                     *
    ' *      *  The shared read-only About property returns information     *
    ' *         about this class.                                           *
    ' *                                                                     *
    ' *      *  The read-only Abstract property returns True for Variants,  *
    ' *         which have no specified type (including either Unknown or   *
    ' *         Null.) This is the case for variants inside arrays.         *
    ' *                                                                     *
    ' *      *  The arraySlice method creates a new variable type based on  *
    ' *         the type of the instance, which has to be an array. The new *
    ' *         is created by removing the toplevel lowerBound and the      *
    ' *         upperBound. The returned variable type will be              *
    ' *         an array of one lower dimension, a variant, or a scalar.    *
    ' *                                                                     *
    ' *      *  The indexed, read-only property BoundSize(d) returns the    *
    ' *         size of an array variable at dimension d (upper bound minus *
    ' *         lower bound plus one.) An error occurs when the variable is *
    ' *         not an array, or the dimension is not defined.              *
    ' *                                                                     *
    ' *      *  The shared method changeArrayBound(fs,d,lu,c) modifies an   *
    ' *         array bound in a fromString expression, returning the       *
    ' *         modified fromString expression:                             *
    ' *                                                                     *
    ' *         + fs: the input fromString expression                       *
    ' *                                                                     *
    ' *         + d: the dimension (rank) to be modified                    *
    ' *                                                                     *
    ' *         + lu: 0 to modify the lowerBound: 1 to modify the upper     *
    ' *           bound                                                     *
    ' *                                                                     *
    ' *         + c: a numeric value to be added to or subtracted from the  *
    ' *           lower or upper bound                                      *
    ' *                                                                     *
    ' *      *  The shared read-only ClassName property returns the name    *
    ' *         "qbVariableType" of this class                              *
    ' *                                                                     *
    ' *      *  The clone method creates a new copy of the object instance  *
    ' *         with identical values, returning this copy as its function  *
    ' *         value.                                                      *
    ' *                                                                     *
    ' *         By default the clone works by (1) converting the clonee     *
    ' *         to its toString representation and (2) assigning this       *
    ' *         string to the clone.                                        *
    ' *                                                                     *
    ' *         However, the optional overload clone(True) will do a faster *
    ' *         clone by copying the state of the clonee to the state of    *
    ' *         the clone. Note that this overload hasn't been fully tested *
    ' *         at this writing while the default clone method has been     *
    ' *         tested.                                                     *
    ' *                                                                     *
    ' *      *  The compareTo(o) method returns True when the type and (for *
    ' *         an array) the number of dimensions and all lower/upper      *
    ' *         bounds of another qbVariableType (o) match the type and     *
    ' *         array information of the object instance.  Otherwise,       *
    ' *         compareTo returns False.                                    *
    ' *                                                                     *
    ' *         Note that compareTo will always return False when the object*
    ' *         instance and/or o are UDTs.                                 *
    ' *                                                                     *
    ' *         The optional overload compareTo(o,e) performs the compari-  *
    ' *         sion, and sets its ByRef string parameter e to an explan-   *
    ' *         ation of why the objects are the same or different.         *
    ' *                                                                     *
    ' *      *  The Shared method containedType(type1, type2) returns True  *
    ' *         when all values in type1 can be converted to valid values of*
    ' *         type2.                                                      *
    ' *                                                                     *
    ' *         The parameters type1 and type2 may be specified as type     *
    ' *         enumerators (of enumerator type ENUvarType) or as objects   *
    ' *         of type qbVariableType.                                     *
    ' *                                                                     *
    ' *         Note that:                                                  *
    ' *                                                                     *
    ' *              + Given the above definition, a type is "contained" in *
    ' *                itself.                                              *
    ' *                                                                     *
    ' *              + The following table shows containment for scalar     *
    ' *                types. Scalar types can be specified as enumerators, *
    ' *                or as qbVariableType objects.                        *
    ' *                                                                     *
    ' *                    boo byt int lng sgl dbl str                      *
    ' *                boo  Y   N   Y   Y   Y   Y   Y                       *
    ' *                byt  N   Y   Y   Y   Y   Y   Y                       *
    ' *                int  N   N   Y   Y   Y   Y   Y                       *
    ' *                lng  N   N   N   Y   Y   Y   Y                       *
    ' *                sgl  N   N   N   N   Y   Y   Y                       *
    ' *                dbl  N   N   N   N   N   Y   Y                       *
    ' *                str  N   N   N   N   N   N   Y                       *
    ' *                                                                     *
    ' *              + Any scalar type is "contained" in the Variant type.  *
    ' *                The Variant type may be specified as an enumerator or*
    ' *                as a qbVariableType although only the latter specifi-*
    ' *                cation provides the type of the value contained in   *
    ' *                the Variant.                                         *
    ' *                                                                     *
    ' *                This is due to our definition, above, of what it is  *
    ' *                to be "contained"; type a is contained in type b when*
    ' *                values of type a can be converted to type b.         *
    ' *                                                                     *
    ' *                Even if a variant V contains an integer, it can be   *
    ' *                assigned a string, therefore the test for containment*
    ' *                in a variant just ignores what is in the variant     *
    ' *                when passed a qbVariableType.                        *
    ' *                                                                     *
    ' *              + The situation is different for arrays and array types*
    ' *                passed to this method must be qbVariableType objects.*
    ' *                                                                     *
    ' *                This method needs to "know" the type inside the      *
    ' *                array and cannot accept ENUvarType values of array.  *
    ' *                                                                     *
    ' *                Each array, of a specific dimensionality, is a type  *
    ' *                in its own right, even when the arrays differ in     *
    ' *                bounds only: for example, array,integer,0,9 is a     *
    ' *                distinct type, with respect to array,integer,1,10.   *
    ' *                                                                     *
    ' *                However, given the above definition, array types can *
    ' *                be contained the one inside the other! The following *
    ' *                rules apply when determining type containment for    *
    ' *                array a1 and array a2:                               *
    ' *                                                                     *
    ' *                - a1 and a2 must have identical numbers of           *
    ' *                  dimensions.                                        *
    ' *                                                                     *
    ' *                - The array scalar type of a1 must be contained in   *
    ' *                  the array scalar type of a2.                       *
    ' *                                                                     *
    ' *                - The size of each of a1's dimensions must be less   *
    ' *                  than or equal to the size of each one of a2's      *
    ' *                  dimensions                                         *
    ' *                                                                     *
    ' *              + No two UDTs are ever contained in each other.        *
    ' *                                                                     *
    ' *      *  The containedTypeWithState(type1, type2) method returns True*
    ' *         when all values in type1 can be converted to valid values of*
    ' *         type2. This method returns the values described under the   *
    ' *         containedType method. It is a non-shared method which on    *
    ' *         first use, creates the type containment table and is there- *
    ' *         fore more efficient than the containedType method.          *
    ' *                                                                     *
    ' *         An overload of this method, containedTypeWithState(type1,   *
    ' *         type2, explanation) places a one-line report in its ByRef   *
    ' *         string parameter, explanation, explaining why type 1 is (or *
    ' *         is not) contained in type 2.                                *
    ' *                                                                     *
    ' *      *  The defaultValue method returns the default value for the   *
    ' *         type (as a .Net value object). The Shared syntax            *
    ' *         defaultValue(e) returns the default value for the type      *
    ' *         identified by the ENUvarType enumerator, e. The syntax      *
    ' *         defaultValue with no operand must be used on a created      *
    ' *         object and it returns the defaultValue for the type in the  *
    ' *         object instance.                                            * 
    ' *                                                                     *
    ' *         For variants, nulls and UDTs this method returns Nothing.   *
    ' *         For arrays the syntax defaultValue with no operand, which   *
    ' *         must be executed with an actual instance, returns the       *
    ' *         array's scalar type. defaultValue(e) cannot be used when e  *
    ' *         is an ENUvarType.vtArray.                                   *
    ' *                                                                     *
    ' *      *  The read-only Dimensions property returns the number of     *
    ' *         dimensions associated with an Array.  This property         *
    ' *         returns 0 with no other error indication when the           *
    ' *         variable is not an Array.                                   *
    ' *                                                                     *
    ' *      *  The dispose method disposes the heap storage associated with*
    ' *         the object (if any) and marks the object not usable.  For   *
    ' *         best results, use this method when finished with the        *
    ' *         instance. If you are truly anal-retentive, see the          *
    ' *         disposeInspect method.                                      *
    ' *                                                                     *
    ' *         The dispose method makes the object not Usable.             *
    ' *                                                                     *
    ' *      *  The disposeInspect method (1) inspects the object for       *
    ' *         internal errors (raising an error if any occur) and (2)     *
    ' *         gets rid of any and all reference variables in the object   *
    ' *         state. For best results, use this method (or the dispose    *
    ' *         method) when finished with the instance.                    *
    ' *                                                                     *
    ' *         The disposeInspect method makes the object not Usable.      *
    ' *                                                                     *
    ' *      *  The fromString method sets the variable type.  This method  *
    ' *         has the following overloads.                                *
    ' *                                                                     *
    ' *         Note: symbols in <corner brackets> represent entry types,   *
    ' *         while unbracketed symbols represent literal entries.        *
    ' *                                                                     *
    ' *         + fromString(<simpleType>) sets the type to one of the      *
    ' *           simple types; <simpleType> should be one of Boolean, Byte,*
    ' *           Integer, Long, Single, Double, String, Unknown or Null.   *
    ' *                                                                     *
    ' *         + fromString("Variant,<varType>") sets the type to Variant, *
    ' *           and the Variant's VarType (the type of the contained      *
    ' *           value) to <varType>.  Note that <varType> can be any      *
    ' *           variable type EXCEPT for Variant or Unknown; in           *
    ' *           particular, variants containing arrays and the Null       *
    ' *           vartype are supported.                                    *
    ' *                                                                     *
    ' *           If <varType> is omitted as in fromString("Variant"), a    *
    ' *           Null variant is created.                                  *
    ' *                                                                     *
    ' *           When the Variant contains an array, <varType> should be a *
    ' *           parenthesized string in fromString's format that defines  *
    ' *           an array as described below.  For example, to define a    *
    ' *           Variant that contains an array of 11 rows and 6 columns,  *
    ' *           use fromString("Variant, (Array, Integer, 0, 10, 0, 5)"). *
    ' *                                                                     *
    ' *           When the Variant contains an array of variants, note that *
    ' *           the type is just Variant. When Variant is the main type,  *
    ' *           it must specify a contained type. But when Variant is     *
    ' *           the array type it cannot specify a contained type. For    *
    ' *           example,                                                  *
    ' *                                                                     *
    ' *           fromString("Variant,(Array,Variant,0,10,0,5)")            *
    ' *                                                                     *
    ' *           specifies a Variant, that contains an array of variants.  *
    ' *                                                                     *
    ' *           See the next subbullet for more explanation of the Array  *
    ' *           specification.                                            *
    ' *                                                                     *
    ' *           The basic idea is that both Variant and Array type speci- *
    ' *           fications have to be in parentheses because they both are *
    ' *           n-tuples of values.                                       *
    ' *                                                                     *
    ' *           A Variant is the ordered pair of the abstract type Variant*
    ' *           and the vartype.                                          *
    ' *                                                                     *
    ' *           An array is the n-tuple of the abstract type array, the   *
    ' *           type of its members (which cannot be an array but can be a*
    ' *           Variant), and the list of its lower and upper bounds.     *
    ' *                                                                     *
    ' *         + fromString("Array,<varType>,<bounds>") sets the type to   *
    ' *           Array, and the type of each entry to varType.  Note that  *
    ' *           varType can be any variable type EXCEPT for Array, Unknown*
    ' *           or Null.  Unlike C, Basic has never supported multiple    *
    ' *           dimensions as arrays of arrays and this is supported by   *
    ' *           explicitly multidimensional arrays.                       *
    ' *                                                                     *
    ' *           If <varType> and <bounds> are omitted as in               *
    ' *           fromString("Array"), the type defaults to an array with   *
    ' *           one entry with lowerBound and upperBound 0, with a Null   *
    ' *           type and value.                                           *
    ' *                                                                     *
    ' *           <bounds> should be the comma-separated list of the lower  *
    ' *           and upper bound of each dimension, from the major to the  *
    ' *           minor dimension.  For example, fromString                 *
    ' *           ("Array,Long,1,10,0,5") defines a Long array of two       *
    ' *           dimensions with 10 columns of 6 rows each with default    *
    ' *           values.                                                   *
    ' *                                                                     *
    ' *           When the array's varType is Variant, it must be specified *
    ' *           abstractly as Variant which makes it possible for the     *
    ' *           variant array to have multiple variant types.             *
    ' *                                                                     *
    ' *         + fromString("UDT,<udtList>") sets the type to User Data    *
    ' *           Type. <udtList> should be a comma-delimited list of       *
    ' *           parenthesized UDT members.                                *
    ' *                                                                     *
    ' *           Each member has the syntax (name,type) where:             *
    ' *                                                                     *
    ' *           - name is the UDT member name: note that this is the only *
    ' *             place where variable names appear in variable types.    *
    ' *                                                                     *
    ' *           - type is the UDT member type expression, and it must be  *
    ' *             the fromString expression for a scalar, a variant, or an*
    ' *             array.                                                  *
    ' *                                                                     *
    ' *           For example take the following UDT.                       *
    ' *                                                                     *
    ' *                Public Type TYPstatus                                *
    ' *                    Dim intValue As Integer                          *
    ' *                    Dim vntValue As Variant                          *
    ' *                    Dim intArray(1 To 10) As Integer                 *
    ' *                    Dim vntArray(1 To 10, 1 To 10) As Variant        *
    ' *                End Type                                             *
    ' *                                                                     *
    ' *           The fromString expression will be:                        *
    ' *                                                                     *
    ' *                UDT,(intValue,Integer),(vntValue,Variant),           *
    ' *                (intArray,Array,Integer,1,10),                       *
    ' *                (vbtArray,Array,Variant,1,10,1,10)                   *
    ' *                                                                     *
    ' *         The formal syntax of fromString style expressions is        *
    ' *         described in the last section of this comment block.        *
    ' *                                                                     *
    ' *      *  The Shared fromString2Typename(fs) method returns the type  *
    ' *         name when provided the fromString expression.               *
    ' *                                                                     *
    ' *      *  The Shared hungarianPrefix(t) method returns the three-digit*
    ' *         Hungarian prefix for the type t, which must be specified    *
    ' *         as an ENUvarType: boo, byt, int, lng, sgl, dbl, str,        *
    ' *         var (variant), arr (array), typ (UDT), unk (unknown) or     *
    ' *         nul (null).                                                 *
    ' *                                                                     *
    ' *      *  The innerType method returns Nothing when the               *
    ' *         type is Unknown, Null or a scalar type. For a Variant or an *
    ' *         array this method returns the handle of the contained       *
    ' *         variant type or the array entry type.                       *
    ' *                                                                     *
    ' *         For a UDT this property must be indexed and it has the      *
    ' *         syntax innerType(x):                                        *
    ' *                                                                     *
    ' *         + The index x may be a number from 1 to the number of       *
    ' *           members in the UDT (available as the UDTmemberCount       *
    ' *           property)                                                 *
    ' *                                                                     *
    ' *         + Alternatively, the index x may be the name of the UDT     *
    ' *           member. Note that udtMemberName(i) can return the member  *
    ' *           name of the indexed inner type.                           *
    ' *                                                                     *
    ' *      *  The inspect(report) method checks the state of the object   *
    ' *         for conformance to rules. Note that inspect does not check  *
    ' *         for user error, but instead for errors resulting from       *
    ' *         software bugs and object misuse.                            *
    ' *                                                                     *
    ' *         In inspect, the report parameter must be a ByRef string, and*
    ' *         it is set to an inspection report.  This object returns True*
    ' *         on success; on failure, it returns False and marks the      *
    ' *         object as not usable.                                       *
    ' *                                                                     *
    ' *         The inspect method applies the following rules.             *
    ' *                                                                     *
    ' *         + The object instance must be usable                        *
    ' *                                                                     *
    ' *         + The type must be compatible with the contained type       *
    ' *           and, when the type is Array, with the bounds:             *
    ' *                                                                     *
    ' *           - If the type is scalar (Boolean, Integer, Long, or       *
    ' *             String) the contained type must be Nothing.             *
    ' *                                                                     *
    ' *           - If the type is Variant the contained type must be Null, *
    ' *             a scalar, or Array.                                     *
    ' *                                                                     *
    ' *           - If the type is Array the contained type must be Null,   *
    ' *             a scalar, or a Variant.                                 *
    ' *                                                                     *
    ' *           - If the type is UDT the contained type must be a collec- *
    ' *             tion of scalar, variant, or array types.                *
    ' *                                                                     *
    ' *         + The type that is contained in the Variant or Array must   *
    ' *           pass its own inspection; each type in a UDT must likewise *
    ' *           pass its own inspection.                                  *
    ' *                                                                     *
    ' *         In addition, and by default, the following rules are also   *
    ' *         checked unless the optional parameter booBasic:=True is     *
    ' *         present.                                                    * 
    ' *                                                                     *
    ' *         + When the object is cloned, the clone must return the same *
    ' *           toString value as the original object.                    *
    ' *                                                                     *
    ' *         + When the fromString value of the object is used to set the*
    ' *           value of a new instance, the compareTo method must indi-  *
    ' *           cate that the original instance and the new instance are  *
    ' *           identical.                                                *
    ' *                                                                     *
    ' *      *  The isArray method returns True (variable type is array     *
    ' *         type) or False (variable type is not array type.)           *
    ' *                                                                     *
    ' *      *  The isomorphicType(type2) method returns True               *
    ' *         when all values in the instance type can be converted to    *
    ' *         valid values of type2 and vice-versa. Type2 must be a       *
    ' *         qbVariableType object.                                      *
    ' *                                                                     *
    ' *         The overload syntax isomorphicType(type2,s) can be used     *
    ' *         to get an explanation of why the types are or are not       *
    ' *         isomorphic. s is a string, passed by reference.             *
    ' *                                                                     *
    ' *      *  The isDefaultValue method tests a .Net value object and     *
    ' *         returns True (.Net value is the default for the variable    *
    ' *         type) or False. The Shared syntax isDefaultValue(n,e) tests *
    ' *         the .Net value object n to see if it matches the default    *
    ' *         value for the type identified by the ENUvarType enumerator, *
    ' *         e. The syntax defaultValue(n) must be used on a created     *
    ' *         object and it tests n against the default for the created   *
    ' *         object's type.                                              *
    ' *                                                                     *
    ' *      *  The isomorphicType(t) method returns True when the          *
    ' *         containedType method returns True for the object instance   *
    ' *         with respect to the qbVariableType t, and for t with re-    *
    ' *         spect to the object instance.                               *
    ' *                                                                     *
    ' *      *  The isScalar method returns True (variable type is of       *
    ' *         scalar type) or False.                                      *
    ' *                                                                     *
    ' *      *  The Shared isScalarType(type) method returns True when type *
    ' *         is one of the Quick Basic types Boolean, Byte, Integer,     *
    ' *         Long, Single, Double or String, False otherwise.            *
    ' *                                                                     *
    ' *         The type may be specified as an ENUtokenType enumerator, or,*
    ' *         as the string name of a type; when it is specified as the   *
    ' *         name of a type case as well as leading/trailing spaces are  *
    ' *         both ignored.                                               *
    ' *                                                                     *
    ' *      *  The isUnknown method returns True (variable type is Unknown)*
    ' *         or False (variable type is known.)                          *
    ' *                                                                     *
    ' *      *  The isUDT method returns True (UDT variable type) or False. *
    ' *                                                                     *
    ' *      *  The isVariant method returns True (variant variable type)   *
    ' *         or False.                                                   *
    ' *                                                                     *
    ' *      *  The indexed and read-write LowerBound(d) property returns   *
    ' *         and can change the lowerbound of an array at the dimension  *
    ' *         d.  d starts at 1 for the major dimension.  If the          *
    ' *         qbVariableType isn't an array or d is invalid an error      *
    ' *         occurs.                                                     *
    ' *                                                                     *
    ' *         Note that the LowerBound may not be changed to a value that *
    ' *         is greater than the UpperBound because this would leave the *
    ' *         qbVariableType object in an invalid state. See redimension  *
    ' *         for a method that can change the lower and/or the upper     *
    ' *         bounds of an array in one statement.                        *
    ' *                                                                     *
    ' *      *  The Shared mkRandomArray method creates a random            *
    ' *         array specifier (as its toString/fromString string in the   *
    ' *         form Array,<type>,<dimensions>), with  anywhere between 1   *
    ' *         and 3 dimensions. Each dimension has random upper and lower *
    ' *         bounds in the range -5..5.                                  *
    ' *                                                                     *
    ' *         This array will have 1..3 random dimensions, with size at   *
    ' *         each dimension restricted to 20 elements, and lower and     *
    ' *         upper bounds in the range -5..5.                            *
    ' *                                                                     *
    ' *         The optional overload syntax mkRandomArray(types) may       *
    ' *         restrict the types to which the array must be restricted,   *
    ' *         and the types may be a blank-separated list of any or all   *
    ' *         of the words Boolean, Byte, Integer, Long, Single, Double,  *
    ' *         String or Variant.                                          *
    ' *                                                                     *
    ' *      *  The Shared mkRandomDomain method returns a random Quick     *
    ' *         Basic type as one of Boolean, Byte, Integer, Long, Single,  *
    ' *         Double, String, Variant, Array, Unknown or Null, and        *
    ' *         returns this type as an ENUvarType enumerator.              *
    ' *                                                                     *
    ' *         The optional overload mkRandomDomain(True) will restrict the*
    ' *         returned domains to the scalar domains Boolean, Byte,       *
    ' *         Integer, Long, Single, Double or String.                    *
    ' *                                                                     *
    ' *      *  The Shared mkRandomScalar method returns a random Quick     *
    ' *         Basic type expressed as the fromString for that type.       *
    ' *         Return values are Boolean, Byte, Integer, Long, Single,     *
    ' *         Double or String.                                           *
    ' *                                                                     *
    ' *         The overload mkRandomScalar(s) may restrict the types       *
    ' *         returned, to the list of blank-delimited type names in s.   *
    ' *                                                                     *
    ' *         See also the mkRandomScalarValue method.                    *
    ' *                                                                     *
    ' *      *  The Shared mkRandomScalarValue method returns the .Net      *
    ' *         version of a random Quick Basic scalar of type (you guessed *
    ' *         it) Boolean, Byte, Integer, Long, Single, Double, or String.*
    ' *                                                                     *
    ' *         The overload mkRandomScalarValue(s) may restrict the types  *
    ' *         of the values returned, to the list of blank-delimited type *
    ' *         names in s.                                                 *
    ' *                                                                     *
    ' *         See also the mkRandomScalar method.                         *
    ' *                                                                     *
    ' *      *  The Shared mkRandomType method returns a random Quick       *
    ' *         Basic type expressed as the fromString for that type.       *
    ' *         This type will be with 20% probability one of array, scalar,*
    ' *         variant, or UDT with 10% probability Unknown and with 10%   *
    ' *         probability Null.                                           *
    ' *                                                                     *
    ' *      *  The Shared mkRandomUDT method creates a random user data    *
    ' *         type specifier (as its toString/fromString string).         *
    ' *                                                                     *
    ' *         This UDT will have 1..10 members. Each member will be a     *
    ' *         random type selected with equal probability from scalar,    *
    ' *         variant, array or user data type with one exception: not    *
    ' *         more than five nested UDTs are allowed.                     *
    ' *                                                                     *
    ' *      *  The Shared mkRandomVariant method creates a random value    *
    ' *         of the Variant type, containing a nonvariant type.          *
    ' *                                                                     *
    ' *         The variant is returned as a String in the form type(value),*
    ' *         where type is one of the Quick Basic scalar types Boolean,  *
    ' *         Byte, Integer, Long, Single, Double, or String and value    *
    ' *         is the value, which will be True, False, a number, or a     *
    ' *         string. This notation is referred to as "decorated"         *
    ' *         notation.                                                   *
    ' *                                                                     *
    ' *      *  The Shared mkRandomVariantValue method creates a random     *
    ' *         Variant type (containing a random scalar type). Note that   *
    ' *         this method can't create a Variant that contains an array.  *
    ' *                                                                     *
    ' *      *  The Shared mkType(e) method creates a new qbVariableType    *
    ' *         based on the ENUvarType enumerator e. The ENUvarType enumer-*
    ' *         ator, e, may not be an Array but it may be a Variant        *
    ' *                                                                     *
    ' *      *  The mkUnusable method marks the object as not usable.       *
    ' *                                                                     *
    ' *      *  The read-write Name property returns and can be set to an   *
    ' *         object instance name.  Name defaults to variableType<nnnn>  *
    ' *         <date> <time> where <nnnn> is a sequence number.            *
    ' *                                                                     *
    ' *      *  The Shared method name2NetType(name) converts the system's  *
    ' *         name for a .Net type (such as System.Int32) to the generic  *
    ' *         name of one of the the .Net types used to support Quick     *
    ' *         Basic variables (such as Integer.) Note that name2NetType   *
    ' *         will convert System.Int64 to Integer.                       *
    ' *                                                                     *
    ' *      *  The Shared method netType2Name(type) converts the generic   *
    ' *         name for a .Net type (such as Integer) to the system        *
    ' *         name of the .Net type (such as System.Int32.)               *
    ' *                                                                     *
    ' *      *  The Shared method netType2QBdomain(value) returns the       *
    ' *         QuickBasic type used to represent the .Net scalar value as  *
    ' *         an ENUvarType.                                              *
    ' *                                                                     *
    ' *         See also qbDomain2NetType.                                  *
    ' *                                                                     *
    ' *      *  The Shared method netValue2QBdomain(value) returns the      *
    ' *         narrowest Quick Basic type to which the .Net object in value*
    ' *         converts without error as an ENUvarType enumerator. This    *
    ' *         method returns ENUvarType.vtUnknown when the .Net object    *
    ' *         does not convert to any Quick Basic type.                   *
    ' *                                                                     *
    ' *         This method won't return Boolean, Array, or Variant.        *
    ' *         Instead, it returns one of Byte, Integer, Long, Single,     *
    ' *         Double, String, Null or Unknown.  Null is returned when     *
    ' *         the .Net value is Nothing; Unknown is returned (with no     *
    ' *         other error indication) when the .Net value converts to no  *
    ' *         other values.                                               * 
    ' *                                                                     *
    ' *         See also netValueInQBdomain.                                *
    ' *                                                                     *
    ' *      *  The Shared method netValue2QBvalue converts a .Net value to *
    ' *         a Quick Basic value, which is returned as its function      *
    ' *         value. This method has the following overloads:             *
    ' *                                                                     *
    ' *         + netValue2QBvalue(value) converts the .Net value to the    *
    ' *           narrowest possible Quick Basic value.                     *
    ' *                                                                     *
    ' *         + netValue2QBvalue(value,type) converts the .Net value to   *
    ' *           the Quick Basic type identified in type. The type may be  *
    ' *           identified as a string or as an ENUvariableType.          *
    ' *                                                                     *
    ' *      *  The netValueInQBdomain method tests .Net scalar values to   *
    ' *         make sure they are in the "domain" (set of values) of the   *
    ' *         Quick Basic variable type. This method has a Shared syntax  *
    ' *         and an "instantiated" syntax.                               *
    ' *                                                                     *
    ' *         The Shared syntax netValueInQBdomain(e,value) tests the .Net*
    ' *         object in value for membership in the Quick Basic type      *
    ' *         identified by the ENUvarType in e.                          * 
    ' *                                                                     *
    ' *         The "instantiated" syntax netValueInQBdomain(value) tests   *
    ' *         the object in value for membership in the Quick Basic type  *
    ' *         in the object instance.                                     * 
    ' *                                                                     *
    ' *         Note that the Object in value must be a .Net scalar. If the *
    ' *         object is not a .Net scalar this method returns False       *
    ' *         with no other error indication unless the Unknown domain is *
    ' *         tested: all .Net values are considered members of this      *
    ' *         domain.                                                     *
    ' *                                                                     *
    ' *         See also netValue2QBdomain.                                 *
    ' *                                                                     *
    ' *      *  The Shared method netValueIsScalar returns True when the    *
    ' *         net value passed is one that can represent a Quick Basic    *
    ' *         nonarray value, eg., one of Boolean, Byte, Short, Integer,  *
    ' *         Single, Double, or String. Otherwise this method returns    *
    ' *         False.                                                      * 
    ' *                                                                     *
    ' *      *  The New method creates the object instance.  Note that the  *
    ' *         VariableType will be Unknown.  To create an object instance,*
    ' *         with another VariableType, use New(<type>), where <type>    *
    ' *         specifies the type in a syntax acceptable to the fromString *
    ' *         method.                                                     *
    ' *                                                                     *
    ' *      *  The Shared method object2Type(o), for any .Net object,      *
    ' *         returns the corresponding fromString expression for its     *
    ' *         type.                                                       *
    ' *                                                                     *
    ' *         If o is a Scalar, with a .Net scalar type that corresponds  *
    ' *         to a Quick Basic type (Boolean, Byte, Short, Integer,       *
    ' *         Single, Double or String), then that QB type is returned.   *
    ' *                                                                     *
    ' *         If o is a .Net long but in the range -2^31..2^31-1 then the *
    ' *         Long QB type is returned.                                   *
    ' *                                                                     *
    ' *         If o is any other value then Unknown is returned.           *
    ' *                                                                     *
    ' *      *  The object2XML method converts the object to eXtended Markup*
    ' *         Language format.                                            *
    ' *                                                                     *
    ' *         This method has the following syntax:                       *
    ' *                                                                     *
    ' *         + object2XML returns the state of the object as a single    *
    ' *           XML tag. A leading comment describes the qbVariableType   *
    ' *           class, the specific variable type, and the current state  *
    ' *           of the cache that is used to save variable type objects   *
    ' *           for reuse, as described below in Cache Considerations.    *
    ' *                                                                     *
    ' *         + object2XML(booIncludeCacheInfo:=False) produces the same  *
    ' *           tag, with one exception: no cache information is included.*
    ' *                                                                     *
    ' *      *  The Shared method qbDomain2NetType(e) converts the type that*
    ' *         is identified by the ENUvarType in e to the .Net type that  *
    ' *         is used to contain values of this type. The e value cannot  *
    ' *         be Unknown, Null or Array. The .Net type is returned as the *
    ' *         string name of the type; it will be in the form SYSTEM.type.*
    ' *                                                                     *
    ' *      *  The redimension(dimension,lower,upper) method redimensions  *
    ' *         an array type at dimension to the lower bound and upper     *
    ' *         bound specified.                                            *
    ' *                                                                     *
    ' *         While redimension doesn't allow lower to be greater than    *
    ' *         upper it does avoid the problem that occurs when you        *
    ' *         need to sequentially change the lower bound to a value that *
    ' *         is higher than the upper bound, and the upper bound to a new*
    ' *         valid value, or vice-versa, using the LowerBound and        *
    ' *         UpperBound properties.                                      *
    ' *                                                                     *
    ' *      *  The scalarDefault method returns the default value applic-  *
    ' *         able to the type:                                           *
    ' *                                                                     *
    ' *         + If the type is Null or Unknown scalarDefault returns      *
    ' *           Nothing                                                   *
    ' *                                                                     *
    ' *         + If the type is scalar, scalarDefault returns the default  *
    ' *           for the scalar type                                       * 
    ' *                                                                     *
    ' *         + If the type is array, scalarDefault returns the default   *
    ' *           for the array entry                                       *
    ' *                                                                     *
    ' *         + If the type is "concrete variant" (a Variant with a known *
    ' *           embedded type) scalarDefault returns the default for the  *
    ' *           embedded type                                             *
    ' *                                                                     *
    ' *         + If the type is "abstract variant" (a Variant with an      *
    ' *           unknown embedded type) scalarDefault returns Nothing      *
    ' *                                                                     *
    ' *      *  The read-only property StorageSpace returns the total number*
    ' *         of abstract cells occupied by the variable type.            *
    ' *                                                                     *
    ' *         Scalars and Variants occupy one cell. Arrays occupy the     *
    ' *         number of cells equivalent to their total size in array     *
    ' *         elements. Unknowns occupy 0 cells, and Nulls occupy 1 cell. *
    ' *         UDTs occupy the sum of space occupied by their members.     *
    ' *                                                                     *
    ' *      *  The Shared string2enuVarType(string) method converts a      *
    ' *         string type name to an enuVarType enumerator.               *
    ' *                                                                     *
    ' *         The string should be one of vtBoolean, vtByte, vtInteger,   *
    ' *         vtLong, vtSingle, vtDouble, vtString, vtVariant, vtArray,   *
    ' *         vtUDT, vtNull or vtUnknown. The prefix vt may be omitted,   *
    ' *         and the name is case-insensitive.                           *
    ' *                                                                     *
    ' *      *  The read-write Tag property may be set to any object to     *
    ' *         associate user data with the object instance. Note that     *
    ' *         the associated object, when it is a reference object, is    *
    ' *         not destroyed when the qbVariableType instance is disposed. *
    ' *                                                                     *
    ' *      *  The test(report) method executes a series of self-tests on  *
    ' *         the object. These tests are not performed on the object     *
    ' *         instance; this would disrupt the object instance for reuse. *  
    ' *                                                                     *
    ' *         Instead, these tests are performed on an internal instance. *
    ' *         The test method uses the fromString method (with toString   *
    ' *         verification in effect) to set the internal instance to each*
    ' *         type including a test array.                                *
    ' *                                                                     *
    ' *         The report parameter is set to a test report.               *
    ' *                                                                     *
    ' *         To suppress the generation of the test method you can set   *
    ' *         the compile-time symbol QBVARIABLETYPETEST_NOTEST to True.  *
    ' *                                                                     *
    ' *      *  The shared, read-only TestAvailable property returns True   *
    ' *         when the test method is supported by the class as compiled, *
    ' *         False otherwise.                                            *
    ' *                                                                     *
    ' *      *  The testEvent(desc,lvlChange) event is fired for a specific *
    ' *         test event, and can be used to display test progress. desc  *
    ' *         is the test description: lvlChange is used to indicate a    *
    ' *         new test level, such that further events represent the      *
    ' *         initiation (or completion) of subgoals, or lvlChange shows  *
    ' *         completion of a goal and return to the next higher goal.    *
    ' *                                                                     *
    ' *         For example, the test method fires a one level message on   *
    ' *         entry. During testing it may fire changes in level, with or *
    ' *         without non-null descriptions, to indicate the initiation or*
    ' *         completion of goals. At the end of a successful test, the   *
    ' *         test method fires testEvent("Test complete", -1) to indicate*
    ' *         completion of the main goal.                                *
    ' *                                                                     *
    ' *         To get this event note that you have to declare the         *
    ' *         qbVariableType object WithEvents. Note that its won't be    *
    ' *         available when the QBVARIABLETYPETEST_NOTEST compile-time   *
    ' *         symbol is True.                                             *
    ' *                                                                     *
    ' *      *  The testProgressEvent(desc, entity, number, count) event is *
    ' *         fired for a specific test event inside a test loop, and this*
    ' *         event can be used to display test progress. In this event:  *
    ' *                                                                     *
    ' *         + desc is the test description                              *
    ' *         + entity is the test entity                                 *
    ' *         + number is the test entity number                          *
    ' *         + count is the total number of entities                     *
    ' *                                                                     *
    ' *         To get this event note that you have to declare the         *
    ' *         qbVariableType object WithEvents. Note that its won't be    *
    ' *         available when the QBVARIABLETYPETEST_NOTEST compile-time   *
    ' *         symbol is True.                                             *
    ' *                                                                     *
    ' *      *  The toContainedName method returns the name of a type       *
    ' *         that is contained in a Variant, an Array, or a UDT. This    *
    ' *         method returns a null string, with no other error           *
    ' *         indication, for a scalar, Unknown or null.                  *
    ' *                                                                     *
    ' *         For a Variant or an Array this method should be specified   *
    ' *         as-is. For a UDT its syntax is toContainedName(x) where x   *
    ' *         is the index of the member (from 1 to UDTmemberCount) or    *
    ' *         the member's name.                                          *
    ' *                                                                     *
    ' *      *  The toDescription method returns a readable description of  *
    ' *         the variable type:                                          *
    ' *                                                                     *
    ' *         + For a scalar one of the following descriptions is         *
    ' *           returned:                                                 *
    ' *                                                                     *
    ' *           - Boolean                                                 *
    ' *           - 8-bit Byte in the range 0..255                          *
    ' *           - 16-bit Integer in the range -32768..32767               *
    ' *           - 32-bit Long integer in the range -2**31..2**31-1        *
    ' *           - Single-precision real number (in .Net's range)          *
    ' *           - Double-precision real number (in .Net's range)          *
    ' *           - String (in .Net's range of string lengths)              *
    ' *           - String (limited to 64K characters)                      *
    ' *                                                                     *
    ' *         + For an abstract Variant "Abstract Variant" is returned    *
    ' *                                                                     *
    ' *         + For a concrete Variant "Variant containing <d>" is        *
    ' *           returned where <d> is the description of the contained    *
    ' *           type                                                      *
    ' *                                                                     *
    ' *         + For an array a string in one of the following forms is    *
    ' *           returned                                                  *
    ' *                                                                     *
    ' *           1-dimensional array containing n elements from m to p:    *
    ' *           each element has the type <typeDescription>               *
    ' *                                                                     *
    ' *           2-dimensional array with n rows (from m to p) and q       *
    ' *           columns (from r to s): each element has the type          *
    ' *           <typeDescription>                                         *
    ' *                                                                     *
    ' *           n-dimensional array with these bounds: list of bounds:    *
    ' *           each element has the type <typeDescription>               *
    ' *                                                                     *
    ' *         + For a UDT, "UDT; list" is returned where list is a list   *
    ' *           of scalar, array and variant descriptions.                *
    ' *                                                                     *
    ' *         + For the unknown data type "Unknown type" is returned      *
    ' *                                                                     *
    ' *         + For the null data type "Null type" is returned            *
    ' *                                                                     *
    ' *      *  The toName method returns just the name of the type; see    *
    ' *         also toString.                                              *
    ' *                                                                     *
    ' *      *  The toString method returns the variable type as a string in*
    ' *         the format acceptable to the fromString method.             *
    ' *                                                                     *
    ' *         If the object is unusable, then toString returns Unusable.  *
    ' *         If an error occurs in the creation of the toString, then    *
    ' *         toString returns error information.                         *
    ' *                                                                     *
    ' *         Note that the toString method, on completion, always        *
    ' *         verifies that the toString and fromString methods are work- *
    ' *         ing correctly, because it will at this time (1) create a new*
    ' *         qbVariableType, (2) set its values using fromString, and    *
    ' *         (3) ensure that it is identical to the object instance,     *
    ' *         using the compareTo method.                                 *
    ' *                                                                     *
    ' *         If the objects differ, a critical error is raised and the   *
    ' *         "Me" object is marked as not usable.                        *
    ' *                                                                     *
    ' *         To suppress this check, use toStringVerify and include its  *
    ' *         optional parameter booVerify:=False in the toStringVerify   *
    ' *         call.                                                       *
    ' *                                                                     *
    ' *      *  The toStringVerify method, as noted above, returns the      *
    ' *         variable type as a string, and supports the booVerify       *
    ' *         optional parameter to suppress the check for correctness    *
    ' *         made by the toString method.                                *
    ' *                                                                     *
    ' *      *  The udtMemberDeref(s) method returns the index assigned to  *
    ' *         the UDT member with name n and can be used only with a UDT  *
    ' *         variable type.                                              *
    ' *                                                                     *
    ' *      *  The read-only UDTmemberCount property returns the count of  *
    ' *         UDT members. If the object doesn't represent a UDT this     *
    ' *         property will raise an error.                               *
    ' *                                                                     *
    ' *      *  The udtMemberName(i) method returns the name of a           *
    ' *         UDT's member (it raises an error when the object instance   *
    ' *         is not a UDT).                                              * 
    ' *                                                                     *
    ' *         The index i must be the integer index from 1 (and up to the *
    ' *         value of UDTmemberCount).                                   *
    ' *                                                                     *
    ' *      *  The indexed and read-write UpperBound(d) property returns   *
    ' *         and can change the upperbound of an array at the dimension  *
    ' *         d.  d starts at 1 for the major dimension.  If the          *
    ' *         qbVariableType isn't an array or d is invalid an error      *
    ' *         occurs.                                                     *
    ' *                                                                     *
    ' *         Note that the UpperBound may not be changed to a value that *
    ' *         is less than the LowerBound because this would leave the    *
    ' *         qbVariableType object in an invalid state. See rebound for  *
    ' *         a method that can change the lower and/or the upper bounds  *
    ' *         of an array in one statement.                               *
    ' *                                                                     *
    ' *      *  The read-only Usable property returns True when the object  *
    ' *         is usable, and False when it is unusable.                   *
    ' *                                                                     *
    ' *      *  The Shared method validFromString(fs) returns True when fs  *
    ' *         is valid as a fromString expression, False otherwise.       *
    ' *                                                                     *
    ' *         This method, to check the fromString, actually creates and  *
    ' *         destroys a qbVariable. It may be useful to access this      *
    ' *         variable after verifying that fs is valid, therefore the    *
    ' *         overload syntax validFromString(fs,v) will return the test  *
    ' *         object in v which must be a qbVariableType.                 *
    ' *                                                                     *
    ' *      *  The read-only VariableType property returns variable type   *
    ' *         as an enumerator of type ENUvarType.  Note that this        *
    ' *         property does not return additional information about       *
    ' *         Variants and Arrays; see instead the VarType, Dimensions,   *
    ' *         LBound, and UBound properties as well as the toString       *
    ' *         method.                                                     *
    ' *                                                                     *
    ' *      *  The read-only, Shared VariableTypes property returns the    *
    ' *         space-delimited list of all the variable types supported:   *
    ' *         "Boolean Byte Integer Long Single Double String Variant     *
    ' *         Array Unknown Null"                                         *
    ' *                                                                     *
    ' *      *  The read-only VarType property returns, for Variant and     *
    ' *         Array variables, the type of the contained value of the     *
    ' *         Variant, or each entry of the Array. For non-variant and    *
    ' *         array variables this property returns Unknown, with no other*
    ' *         error indication.                                           *
    ' *                                                                     *
    ' *      *  The Shared method vtPrefixAdd adds the prefix vt to a       *
    ' *         variable type when it isn't already present. Note that the  *
    ' *         canonical ENUvarType names are defined with a name that     *
    ' *         starts with vt, and this method ensures that prefix is      *
    ' *         present.                                                    *
    ' *                                                                     *
    ' *      *  The Shared method vtPrefixRemove removes the prefix vt from *
    ' *         variable type when it is present. Note that the             *
    ' *         canonical ENUvarType names are defined with a name that     *
    ' *         starts with vt, and this method removes the prefix.         *
    ' *                                                                     *
    ' *                                                                     *
    ' * USE AS A SHARED TOOL SOURCE --------------------------------------- *
    ' *                                                                     *
    ' * Consider merely Importing this class, or merely allocating a Shared,*
    ' * code-only instance of this object if all you need are these Shared  *
    ' * properties and methods:               			                    *
    ' *                                                                     *
    ' *      *  About: return about information                             *
    ' *                                                                     *
    ' *      *  ClassName: return the name of this class                    *
    ' *                                                                     *
    ' *      *  changeArrayBound: modifies array bounds                     *
    ' *                                                                     *
    ' *      *  containedType: test type containment where the types are    *
    ' *         specified as parameters                                     *
    ' *                                                                     *
    ' *      *  defaultValue: return default value when the type is         *
    ' *         specified as a parameter                                    *
    ' *                                                                     *
    ' *      *  fromString2TypeName(fs): extracts the type name             *
    ' *                                                                     *
    ' *      *  hungarianPrefix(t) returns the Hungarian prefix for the type*
    ' *                                                                     *
    ' *      *  isDefaultValue: tests .Net object against default value for *
    ' *         a type                                                      *
    ' *                                                                     *
    ' *      *  isScalarType: tell caller whether type is one of the scalar *
    ' *         Quick Basic types                                           *
    ' *                                                                     *
    ' *      *  mkRandomArray: create a random array as a fromString        *
    ' *                                                                     *
    ' *      *  mkRandomScalar: create a random Quick Basic scalar as a     *
    ' *         fromString                                                  *
    ' *                                                                     *
    ' *      *  mkRandomType: create a random Quick Basic type as a         *
    ' *         fromString                                                  *
    ' *                                                                     *
    ' *      *  mkRandomUDT: create a random Quick Basic user data type as  *
    ' *         a fromString.                                               *
    ' *                                                                     *
    ' *      *  mkRandomVariant: create a random Quick Basic Variant type as*
    ' *         a fromString.                                               *
    ' *                                                                     *
    ' *      *  mkRandomVariantValue: create a random Quick Basic Variant   *
    ' *         value as a decorated string.                                *
    ' *                                                                     *
    ' *      *  mkType: create a qbVariableType                             *
    ' *                                                                     *
    ' *      *  name2NetType: map the .Net system name to the generic name, *
    ' *         of a .Net type that is used to support a Quick Basic type   *
    ' *                                                                     *
    ' *      *  netType2Name: map the generic name, of a .Net type that is  *
    ' *         used to support a Quick Basic type to the .Net system name  * 
    ' *                                                                     *
    ' *      *  netType2QBdomain: map the .Net type to the Quick Basic type *
    ' *                                                                     *
    ' *      *  netValue2QBdomain: map .Net value to the Quick Basic type   *
    ' *                                                                     *
    ' *      *  netValue2QBvalue: map the .Net value to the Quick Basic     *
    ' *         value                                                       *
    ' *                                                                     *
    ' *      *  netValueIsScalar: tell the caller if the .Net value is a    *
    ' *         scalar                                                      *
    ' *                                                                     *
    ' *      *  netValueInQBdomain: Test .Net value to make sure it is in   *
    ' *         the Quick Basic domain                                      *
    ' *                                                                     *
    ' *      *  object2Type: convert .Net object to a fromString expression *
    ' *                                                                     *
    ' *      *  qbDomain2NetType: convert the quickBasic domain to a .Net   *
    ' *         type                                                        *
    ' *                                                                     *
    ' *      *  randomQBtype: return a random Quick Basic type              *
    ' *                                                                     *
    ' *      *  string2enuVarType: convert the string to a variant type     *
    ' *         enumerator                                                  *
    ' *                                                                     *
    ' *      *  TestAvailable: report whether this object version exposes a *
    ' *         test method                                                 *
    ' *                                                                     *
    ' *      *  validFromString: check the fromString for validity          *
    ' *                                                                     *
    ' *      *  VariableTypeList: return the supported variable types       *
    ' *                                                                     *
    ' *      *  vtPrefixAdd: conditionally add the variable type prefix     *
    ' *                                                                     *
    ' *      *  vtPrefixRemove: conditionally remove the variable type      *
    ' *         prefix                                                      *
    ' *                                                                     *
    ' *                                                                     *
    ' * COMPILE-TIME SYMBOLS ---------------------------------------------- *
    ' *                                                                     *
    ' * Set the compile-time symbol QBVARIABLETYPE_NOTEST to True in the    *
    ' * project to suppress the generation of the test method.              *
    ' *                                                                     *
    ' *                                                                     *
    ' * TYPE SPECIFICATION FOR FROMSTRING/TOSTRING: BNF ------------------- *
    ' *                                                                     *
    ' * The following Backus-Naur Form syntax describes the syntax of the   *
    ' * first parameter of the fromString method, and the format of the     *
    ' * strings returned by the toString method.                            *
    ' *                                                                     *
    ' *                                                                     *
    ' *      typeSpecification := baseType | udt                            *
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
    ' *                                                                     *
    ' *                                                                     *
    ' * MULTITHREADING CONSIDERATIONS ------------------------------------- *
    ' *                                                                     *
    ' * Multiple instances of this class may simultaneously execute in      *
    ' * different threads. However, the same instance may not run code in   *
    ' * multiple threads.                                                   *
    ' *                                                                     *
    ' *                                                                     *
    ' * CACHE CONSIDERATIONS ---------------------------------------------- *
    ' *                                                                     *
    ' * This object avoids excessive parsing of "fromString" expressions to *
    ' * create types by using a "cache" to save parsed types.               *
    ' *                                                                     *
    ' * The cache is a keyed Collection in Shared storage. Each Item of this*
    ' * collection contains the clone of a pre-existing variable type       *
    ' * object; the key of each Item is its fromString expression.          *
    ' *                                                                     *
    ' * The cache will contain a maximum of 100 entries and the oldest      *
    ' * entries are dropped when the cache is full.                         *
    ' *                                                                     *
    ' * Information about the cache, in the form of a list of its entries   *
    ' * and its maximum size, is available by default when the state of any *
    ' * instance is displayed using object2XML: see the object2XML method   *
    ' * for more information.                                               *
    ' *                                                                     *
    ' ***********************************************************************

    ' ***********************************************************************
    ' *                                                                     *
    ' * "In whatsoever mode, or by whatsoever means, our knowledge may      *
    ' *  relate to objects, it is at least quite clear, that the only matter*
    ' *  it relates to them, is by means of an intuition."                  *
    ' *                                                                     *
    ' *                        - Kant, Critique of Pure Reason              *
    ' *                                                                     *
    ' ***********************************************************************

    ' ***** Shared *****
    ' --- Utilities
    Private Shared _OBJutilities As utilities.utilities
    ' --- Object sequence number
    Private Shared _INTsequence As Integer
    ' --- Cache
    Private Enum _ENUcacheExistence
        notAvailable
        beingCreated
        available
    End Enum
    Private Shared _INTcacheExistence As Integer    ' 1: cache object exists, else 0
    Private Class _cache
        Public intCacheLimit As Integer             ' Maximum size
        Public colCache As Collection               ' Of qbVariableType, contains
                                                    ' clones: key is fromString
                                                    ' expression
    End Class
    Private Shared _OBJcache As _cache

    ' ***** Collection handling *****
    Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

    ' ***** Variable types *****
    ' --- List of types
    Private Const SCALARTYPES As String = _
    "Boolean Byte Integer Long Single Double String"
    Private Const VARIABLETYPES As String = _
    SCALARTYPES & " Variant Array UDT Null Unknown"
    ' --- Variable types as enumerator  
    Public Enum ENUvarType
        vtBoolean           ' Boolean
        vtByte              ' 8 bit unsigned integer
        vtInteger           ' 16 bit signed integer
        vtLong              ' 32 bit signed integer (VB.Net Long not supported)
        vtSingle            ' Single precision (uses VB.Net singles, does not accurately
                            ' reflect QuickBasic precision)
        vtDouble            ' Double precision (uses VB.Net doubles, does not accurately
                            ' reflect QuickBasic precision)
        vtString            ' Variable length string
        vtVariant           ' Object 
        vtArray             ' Array
        vtUDT               ' User-defined type
        vtUnknown           ' Default 
        vtNull              ' Null: used when the VarType is requested of a noncontainer 
    End Enum
    ' --- Default values
    Private Const DEFAULT_BOOLEAN As Boolean = False
    Private Const DEFAULT_BYTE As Byte = 0
    Private Const DEFAULT_INTEGER As Short = 0
    Private Const DEFAULT_LONG As Integer = 0
    Private Const DEFAULT_SINGLE As Single = 0
    Private Const DEFAULT_DOUBLE As Double = 0
    Private Const DEFAULT_STRING As String = ""
    Private Const DEFAULT_VARIANT As Object = Nothing
    Private Const DEFAULT_UDT As Object = Nothing

    ' ***** State *****
    Friend Structure TYPstate
        Dim booUsable As Boolean           ' Object usability
        Dim strName As String              ' Instance name
        Dim enuVariableType As ENUvarType  ' Main type
        Dim objVarType As Object           ' Contained type(s):
                                           ' Nothing (for a scalar, Unknown or null)
                                           ' Contained type for Variant
                                           ' Entry type for Array
                                           ' Collection for UDT:
                                           ' key is member name:
                                           ' data is 3-member subcollection:
                                           ' Item(1) is member index:
                                           ' Item(2) is member name:
                                           ' Item(3) is member as a qbVariableType
        Dim colBounds As Collection        ' No key: data is 2-element collection:
                                           ' item(1): lowerBound
                                           ' item(2): upperBound
        Dim colTypeOrdering As Collection  ' Type ordering
        Dim booContained(,) As Boolean     ' Type containment
        Dim objTag As Object               ' User object  
    End Structure
    Private USRstate As TYPstate

    ' ***** Constants *****
    ' --- Name of class
    Private Const CLASS_NAME As String = "qbVariableType"
    ' --- Inspection rules
    Private Const INSPECTION_USABLE As String = "Object instance must be usable"
    Private Const INSPECTION_TYPECOMPATIBLE As String = _
            "Type must be compatible with contained value and/or bounds"
    Private Const INSPECTION_CONTAINEDOK As String = _
            "Contained variable type(s) must pass their own inspection"
    Private Const INSPECTION_CLONEOK As String = _
            "Clone returns identical toString value to original object"
    Private Const INSPECTION_TOFROMSTRING As String = _
            "toString used as fromString produces identical objects"
    ' --- Basic type descriptions
    Private Const TODESCRIPTION_BOOLEAN As String = "Boolean"
    Private Const TODESCRIPTION_BYTE As String = "8-bit Byte in the range 0..255"
    Private Const TODESCRIPTION_INTEGER As String = "16-bit Integer in the range -32768..32767"
    Private Const TODESCRIPTION_LONG As String = "32-bit Long integer in the range -2**31..2**31-1"
    Private Const TODESCRIPTION_SINGLE As String = "Single-precision real number (in .Net's range)"
    Private Const TODESCRIPTION_DOUBLE As String = "Double-precision real number (in .Net's range)"
    Private Const TODESCRIPTION_STRING As String = "String (limited to 64K characters)"
    Private Const TODESCRIPTION_STRINGNOLIMIT As String = "String (in .Net's range of string lengths)"
    Private Const TODESCRIPTION_ABSTRACTVARIANT As String = "Abstract Variant"
#If Not QBVARIABLETEST_NOTEST Then
        ' --- Testing            
        ' Fromstrings for test method
        Private Const TEST_FROMSTRING_LIST_BASE As String = _
                "Unknown" & vbNewLine & _
                "Boolean" & vbNewLine & _
                "Byte" & vbNewLine & _
                "Integer" & vbNewLine & _
                "Long" & vbNewLine & _
                "Single" & vbNewLine & _
                "Double" & vbNewLine & _
                "String" & vbNewLine & _
                "Array,Boolean,0,0" & vbNewLine & _
                "Array,Byte,0,1" & vbNewLine & _
                "Array,Integer,0,2" & vbNewLine & _
                "Array,Long,0,3" & vbNewLine & _
                "Array,Single,0,0,0,0" & vbNewLine & _
                "vtArray,vtSingle,0,0,0,0" & vbNewLine & _
                "Array,Double,0,0,0,0,0,0" & vbNewLine & _
                "Array,String,0,9,5,10,-5,5" & vbNewLine & _
                "Array,Variant,0,0" & vbNewLine & _
                "Variant,Null" & vbNewLine & _
                "Variant,Boolean" & vbNewLine & _
                "Variant,Byte" & vbNewLine & _
                "Variant,Integer" & vbNewLine & _
                "Variant,Long" & vbNewLine & _
                "Variant,Single" & vbNewLine & _
                "Variant,Double" & vbNewLine & _
                "Variant,vtDouble" & vbNewLine & _
                "Variant,String" & vbNewLine & _
                "Variant,(Array,String,0,9,5,10,-5,5)" & vbNewLine & _
                "Variant,(Array,Variant,0,9,5,10,-5,5)" & vbNewLine & _
                "vtVariant,(vtArray,vtString,0,9,5,10,-5,5)" & vbNewLine & _
                "UDT," & _
                "(intValue,Integer)," & _
                "(vntValue,Variant)," & _
                "(intArray,Array,Integer,1,10)," & _
                "(vbtArray,Array,Variant,1,10,1,10)" & vbNewLine & _
                "array," & _
                "(udt," & _
                "(member01,byte)," & _
                "(member02,array,integer,1,10)," & _
                "(member03,variant,(array,(udt,(member01,byte)),1,10)))," & _
                "-5,10,1,3"
        ' Test containment list
        Private Const TEST_CONTAINMENT As String = _
            "Boolean:Boolean:True " & vbNewLine & _
            "Boolean:Byte:   False" & vbNewLine & _
            "Boolean:Integer:True " & vbNewLine & _
            "Boolean:Long:   True " & vbNewLine & _
            "Boolean:Single: True " & vbNewLine & _
            "Boolean:Double: True " & vbNewLine & _
            "Boolean:String: True " & vbNewLine & _
            "Byte:   Boolean:False" & vbNewLine & _
            "Byte:   Byte:   True " & vbNewLine & _
            "Byte:   Integer:True " & vbNewLine & _
            "Byte:   Long:   True " & vbNewLine & _
            "Byte:   Single: True " & vbNewLine & _
            "Byte:   Double: True " & vbNewLine & _
            "Byte:   String: True " & vbNewLine & _
            "Integer:Boolean:False" & vbNewLine & _
            "Integer:Byte:   False" & vbNewLine & _
            "Integer:Integer:True " & vbNewLine & _
            "Integer:Long:   True " & vbNewLine & _
            "Integer:Single: True " & vbNewLine & _
            "Integer:Double: True " & vbNewLine & _
            "Integer:String: True " & vbNewLine & _
            "Long:   Boolean:False" & vbNewLine & _
            "Long:   Byte:   False" & vbNewLine & _
            "Long:   Integer:False" & vbNewLine & _
            "Long:   Long:   True " & vbNewLine & _
            "Long:   Single: True " & vbNewLine & _
            "Long:   Double: True " & vbNewLine & _
            "Long:   String: True " & vbNewLine & _
            "Single: Boolean:False" & vbNewLine & _
            "Single: Byte:   False" & vbNewLine & _
            "Single: Integer:False" & vbNewLine & _
            "Single: Long:   False" & vbNewLine & _
            "Single: Single: True " & vbNewLine & _
            "Single: Double: True " & vbNewLine & _
            "Single: String: True " & vbNewLine & _
            "Double: Boolean:False" & vbNewLine & _
            "Double: Byte:   False" & vbNewLine & _
            "Double: Integer:False" & vbNewLine & _
            "Double: Long:   False" & vbNewLine & _
            "Double: Single: False" & vbNewLine & _
            "Double: Double: True " & vbNewLine & _
            "Double: String: True " & vbNewLine & _
            "String: Boolean:False" & vbNewLine & _
            "String: Byte:   False" & vbNewLine & _
            "String: Integer:False" & vbNewLine & _
            "String: Long:   False" & vbNewLine & _
            "String: Single: False" & vbNewLine & _
            "String: Double: False" & vbNewLine & _
            "String: String: True " & vbNewLine & _
            "String: Boolean:False" & vbNewLine & _
            "Variant:Boolean:False" & vbNewLine & _
            "Variant:Byte:   False" & vbNewLine & _
            "Variant:Integer:False" & vbNewLine & _
            "Variant:Single: False" & vbNewLine & _
            "Variant:Double: False" & vbNewLine & _
            "Variant:String: False" & vbNewLine & _
            "Variant:Variant:True " & vbNewLine & _
            "Boolean:Variant:True " & vbNewLine & _
            "Byte:   Variant:True " & vbNewLine & _
            "Integer:Variant:True " & vbNewLine & _
            "Long:   Variant:True " & vbNewLine & _
            "Single: Variant:True " & vbNewLine & _
            "Double: Variant:True " & vbNewLine & _
            "String: Variant:True " & vbNewLine & _
            "Variant:Variant:True " & vbNewLine & _
            "Array:  Boolean:False" & vbNewLine & _
            "Array:  Byte:   False" & vbNewLine & _
            "Array:  Integer:False" & vbNewLine & _
            "Array:  Single: False" & vbNewLine & _
            "Array:  Double: False" & vbNewLine & _
            "Array:  String: False" & vbNewLine & _
            "Boolean:Array:  False" & vbNewLine & _
            "Byte:   Array:  False" & vbNewLine & _
            "Integer:Array:  False" & vbNewLine & _
            "Long:   Array:  False" & vbNewLine & _
            "Single: Array:  False" & vbNewLine & _
            "Double: Array:  False" & vbNewLine & _
            "String: Array:  False" & vbNewLine & _
            "Variant:Array:  False"
        ' Testing the defaultValue method
        Private Const TEST_DEFAULTVALUE As String = _
            "Boolean,Boolean(""False"")" & vbNewLine & _
            "Byte,   Byte(0)" & vbNewLine & _
            "Integer,Short(0)" & vbNewLine & _
            "Long,   Integer(0)" & vbNewLine & _
            "Single, Single(0)" & vbNewLine & _
            "Double, Double(0)" & vbNewLine & _
            "String, String("""")" & vbNewLine & _
            "Variant,Nothing"
#End If
    ' --- About information Easter egg
    Private Const ABOUTINFO As String = _
        "variableType" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "This class represents the type of a quickBasicEngine variable, " & _
        "including support for an unknown type and Shared methods for " & _
        "relating .Net types to Quick Basic types." & vbNewLine & vbNewLine & _
        "This class was developed commencing on 4/5/2003 by" & vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    ' --- Narrow-wide scalar type relations for containedType
    Private Const TYPEORDERING As String = _
    "BOOLEAN BYTE INTEGER LONG SINGLE DOUBLE STRING"
    Private Const TYPECONTAINMENT As String = _
    "Y N Y Y Y Y Y N Y Y Y Y Y Y N N Y Y Y Y Y N N N Y Y Y Y N N N N Y Y Y N N N N N Y Y N N N N N N Y"
    ' --- Cache size
    Private Const DEFAULT_CACHE_LIMIT As Integer = 100

    ' ***** Type requirement *****
    Private Enum ENUcomplexTypeRequirement
        variantType                 ' Any variant
        scalarVariantType           ' Any non-array variant
        arrayType                   ' Array 
        unknown                     ' Not known
    End Enum

#If Not QBVARIABLETEST_NOTEST Then
        ' ***** Public events ***************************************************
        ' --- Test event
        Public Event testEvent(ByVal strDesc As String, ByVal intLevel As Integer)
        ' --- Test progress report
        Public Event testProgressEvent(ByVal strDesc As String, _
                                       ByVal strEntity As String, _
                                       ByVal intEntityNumber As Integer, _
                                       ByVal intEntityCount As Integer)
#End If

    ' ***** Explanation from compareTo *****
    Private STRcompareToExplanation As String

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
            Return (ABOUTINFO)
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Return True for abstract variants only
    '
    '
    Public ReadOnly Property Abstract() As Boolean
        Get
            If Not checkUsable_("Abstract Get") Then Return (False)
            With USRstate
                If .enuVariableType <> ENUvarType.vtVariant Then
                    Return (False)
                End If
                Return (.objVarType Is Nothing)
            End With
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Slice and dice
    '
    '
    Public Function arraySlice() As qbVariableType
        If Not checkUsable_("arraySlice") Then Return (Nothing)
        With Me
            If Not .isArray Then
                Return (mkType(ENUvarType.vtUnknown))
            End If
            ' --- Make the fromString
            Dim strFromstring As String = .VarType.ToString
            If InStr(strFromstring, ",") <> 0 Then
                strFromstring = "(" & strFromstring & ")"
            End If
            If .Dimensions > 1 Then
                strFromstring = "Array," & strFromstring
                Dim intIndex1 As Integer
                For intIndex1 = 2 To .Dimensions
                    _OBJutilities.append(strFromstring, _
                                         ",", _
                                         .LowerBound(intIndex1) & _
                                         "," & _
                                         .UpperBound(intIndex1))
                Next intIndex1
            End If
            ' --- Make the slice object
            Dim objSlice As qbVariableType
            Try
                objSlice = New qbVariableType
                objSlice.fromString(strFromstring)
            Catch
                errorHandler_("Cannot create qbVariableType: " & _
                            Err.Number & " " & Err.Description, _
                            "arraySlice", _
                            "Returning Nothing")
                Return (Nothing)
            End Try
            Return (objSlice)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return array size at a specified dimension
    '
    '
    Public ReadOnly Property BoundSize(ByVal intDimension As Integer) As Integer
        Get
            If Not checkUsable_("BoundSize get") Then Return (0)
            With Me
                Return (.UpperBound(intDimension) - .LowerBound(intDimension) + 1)
            End With
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Modify array bounds
    '
    '
    Public Shared Function changeArrayBound(ByVal strFromstring As String, _
                                            ByVal intRank As Integer, _
                                            ByVal intLU As Integer, _
                                            ByVal intChange As Integer) As String
        Dim objVT As qbVariableType
        If Not validFromString(strFromstring, objVT) Then
            errorHandler_("Invalid fromString expression " & _
                          _OBJutilities.enquote(strFromstring), _
                          ClassName, "changeArrayBound", _
                          "Returning unchanged fromString expression")
            Return strFromstring
        End If
        With objVT
            If Not .isArray Then
                errorHandler_("fromString does not represent an array", _
                              ClassName, "changeArrayBound", _
                              "Returning unchanged fromString expression")
                Return strFromstring
            End If
            If intRank < 1 OrElse intRank > .Dimensions Then
                errorHandler_("Dimension " & intRank & " is not valid", _
                              ClassName, "changeArrayBound", _
                              "Returning unchanged fromString expression")
                Return strFromstring
            End If
            If intLU <> 0 OrElse intRank > .Dimensions Then
                errorHandler_("Dimension " & intRank & " is not valid", _
                              ClassName, "changeArrayBound", _
                              "Returning unchanged fromString expression")
                Return strFromstring
            End If
            .redimension(intRank, _
                         CInt(IIf(intLU = 0, .LowerBound(intRank) + intChange, .LowerBound(intRank))), _
                         CInt(IIf(intLU = 1, .LowerBound(intRank), .UpperBound(intRank) + intChange)))
            Dim strReturn As String = .ToString
            .dispose()
            Return strReturn
        End With
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
    ' Create and return a clone
    '
    '
    ' --- Enforces strict type and usability.
    Public Function clone() As qbVariableType
        If Not checkUsable_("clone") Then Return (Nothing)
        Return (CType(clone_(), qbVariableType))
    End Function
    ' --- Enforces strict type and usability: does a state clone
    Public Function clone(ByVal booStateClone As Boolean) As qbVariableType
        If Not checkUsable_("clone") Then Return (Nothing)
        If Not booStateClone Then Return CType(clone_(), qbVariableType)
        Dim objClone As qbVariableType
        Try
            objClone = New qbVariableType(True)
        Catch
            errorHandler_("Cannot clone for a statewise clone: " & _
                          Err.Number & " " & Err.Description, _
                          "clone", _
                          "Making object unusable: returning Nothing")
            Me.mkUnusable() : Return Nothing
        End Try
        objClone.State = Me.State
        Return objClone
    End Function
    ' --- Implements cloning
    Private Function clone_() As Object Implements ICloneable.Clone
        Dim objClone As qbVariableType
        Try
            objClone = New qbVariableType
        Catch
            errorHandler_("Cannot create the clone object: " & _
                          Err.Number & " " & Err.Description, _
                          "clone", _
                          "Marking object unusable and returning Nothing")
            Me.mkUnusable() : Return (Nothing)
        End Try
        If Not objClone.fromString(Me.ToString) Then
            errorHandler_("Cannot assign values to the clone object: " & _
                          Err.Number & " " & Err.Description, _
                          "clone", _
                          "Marking object unusable and returning Nothing")
            objClone.dispose()
            Me.mkUnusable()
            Return (Nothing)
        End If
        objClone.Name = "Obj#" & Mid(objClone.Name, 15, 4) & ": cloneOf[" & Me.Name & "]"
        Return (objClone)
    End Function

    ' ----------------------------------------------------------------------
    ' Return -1 (object instances have the same type) or 0
    '
    '
    ' --- Enforces strict type and usability
    ' No explanation provided
    Public Overloads Function compareTo(ByVal objQBvariableType As qbVariableType) _
           As Boolean
        Dim strExplanation As String
        Return (Me.compareTo(objQBvariableType, strExplanation))
    End Function
    ' Explanation provided
    Public Overloads Function compareTo(ByVal objQBvariableType As qbVariableType, _
                                        ByRef strExplanation As String) _
           As Boolean
        If Not checkUsable_("compareTo") Then Return (False)
        If Not objQBvariableType.Usable Then
            errorHandler_("Compared qbVariableType isn't usable", _
                          "compareTo", _
                          "Returning False")
            Return (False)
        End If
        Dim intCompareTo As Integer = compareTo_(objQBvariableType)
        strExplanation = STRcompareToExplanation
        Select Case intCompareTo
            Case 0 : Return (False)
            Case -1 : Return (True)
            Case Else
                errorHandler_("Programming error: " & _
                              "Unexpected value " & intCompareTo & " " & _
                              "returned from CompareTo implementation", _
                              "compareTo_", _
                              "Marking the object instance not usable and " & _
                              "returning False")
                Me.mkUnusable() : Return (False)
        End Select
    End Function
    ' --- Implements the operation
    Private Function compareTo_(ByVal objObject As Object) _
            As Integer Implements IComparable.CompareTo
        STRcompareToExplanation = ""
        If (Me Is objObject) Then
            STRcompareToExplanation = "The instance object and the object passed in the operand " & _
                                      "are the same object"
            Return (-1)
        End If
        With CType(objObject, qbVariableType)
            If Me.VariableType <> .VariableType Then
                STRcompareToExplanation = "The main qb variable type in the instance object is " & _
                                          Me.VariableType.ToString & " while " & _
                                          "the main qb type in the operand object is " & _
                                          .VariableType.ToString
                Return (0)
            End If
            If Me.isUDT OrElse .isUDT Then
                STRcompareToExplanation = "The object instance and/or the compared object " & _
                                          "are UDTs"
                Return (0)
            End If
            If Me.VarType <> .VarType Then
                STRcompareToExplanation = "The contained qb variable type in the instance object is " & _
                                          Me.VarType.ToString & " while " & _
                                          "the contained qb type in the operand object is " & _
                                          .VarType.ToString
                Return (0)
            End If
            If Me.Dimensions <> .Dimensions Then
                STRcompareToExplanation = "The instance object has " & _
                                          Me.Dimensions & " dimension(s) while " & _
                                          "the operand object has " & _
                                          .Dimensions & " dimension(s)"
                Return (0)
            End If
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Dimensions
                If Me.LowerBound(intIndex1) <> .LowerBound(intIndex1) Then
                    STRcompareToExplanation = "The instance object has the lower bound " & _
                                              Me.LowerBound(intIndex1) & " " & _
                                              "at dimension " & intIndex1 & " while " & _
                                              "the operand object has lower bound " & _
                                              .LowerBound(intIndex1)
                    Return (0)
                End If
                If Me.UpperBound(intIndex1) <> .UpperBound(intIndex1) Then
                    STRcompareToExplanation = "The instance object has the upper bound " & _
                                              Me.UpperBound(intIndex1) & " " & _
                                              "at dimension " & intIndex1 & " while " & _
                                              "the operand object has the upper bound " & _
                                              .UpperBound(intIndex1)
                    Return (0)
                End If
            Next intIndex1
            STRcompareToExplanation = "The instance object and the operand object describe " & _
                                      "identical Quick Basic data types"
            Return (-1)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Test whether type 1 is narrower than, or identical to, type 2
    '
    '
    ' --- Shared versions
    ' Shared version using enumerators
    Public Shared Function containedType(ByVal enuType1 As ENUvarType, _
                                         ByVal enuType2 As ENUvarType) _
           As Boolean
        Dim objType(1) As qbVariableType
        objType(0) = mkType(enuType1) : objType(1) = mkType(enuType2)
        If (objType(0) Is Nothing) OrElse (objType(1) Is Nothing) Then
            Return (False)
        End If
        Dim booResult As Boolean = containedType(objType(0), objType(1))
        objType(0).dispose() : objType(0) = Nothing
        objType(1).dispose() : objType(1) = Nothing
        Return (booResult)
    End Function
    ' Shared version using enumerator and object
    Public Shared Function containedType(ByVal enuType1 As ENUvarType, _
                                         ByVal objType2 As qbVariableType) _
           As Boolean
        Dim objType As qbVariableType
        objType = mkType(enuType1)
        If (objType Is Nothing) Then Return (False)
        Dim booResult As Boolean = containedType(objType, objType2)
        objType.dispose() : objType = Nothing
        Return (booResult)
    End Function
    ' Shared version using object and enumerator
    Public Shared Function containedType(ByVal objType1 As qbVariableType, _
                                         ByVal enuType2 As ENUvarType) _
           As Boolean
        Dim objType As qbVariableType
        objType = mkType(enuType2)
        If (objType Is Nothing) Then Return (False)
        Dim booResult As Boolean = containedType(objType1, objType)
        objType.dispose() : objType = Nothing
        Return (booResult)
    End Function
    ' Shared version using objects
    Public Shared Function containedType(ByVal objType1 As qbVariableType, _
                                         ByVal objType2 As qbVariableType) _
           As Boolean
        Dim booContained(,) As Boolean
        Dim colTypeOrdering As Collection
        Dim strExplanation As String
        Return (containedType_(objType1, _
                              objType2, _
                              colTypeOrdering, _
                              booContained, _
                              strExplanation))
    End Function
    ' --- Versions with state
    ' Unshared version using enumerators
    Public Function containedTypeWithState(ByVal enuType1 As ENUvarType, _
                                           ByVal enuType2 As ENUvarType) _
           As Boolean
        If Not checkUsable_("containedTypeWithState") Then Return (False)
        Dim objType(1) As qbVariableType
        objType(0) = mkType(enuType1) : objType(1) = mkType(enuType2)
        If (objType(0) Is Nothing) OrElse (objType(1) Is Nothing) Then
            Return (False)
        End If
        Dim booResult As Boolean = containedTypeWithState(objType(0), objType(1))
        objType(0).dispose() : objType(0) = Nothing
        objType(1).dispose() : objType(1) = Nothing
        Return (booResult)
    End Function
    ' Unshared version using enumerator and object
    Public Function containedTypeWithState(ByVal enuType1 As ENUvarType, _
                                           ByVal objType2 As qbVariableType) _
           As Boolean
        If Not checkUsable_("containedTypeWithState") Then Return (False)
        Dim objType As qbVariableType
        objType = mkType(enuType1)
        If (objType Is Nothing) Then Return (False)
        Dim booResult As Boolean = containedTypeWithState(objType, objType2)
        objType.dispose() : objType = Nothing
        Return (booResult)
    End Function
    ' Unshared version using object and enumerator
    Public Function containedTypeWithState(ByVal objType1 As qbVariableType, _
                                           ByVal enuType2 As ENUvarType) _
           As Boolean
        If Not checkUsable_("containedTypeWithState") Then Return (False)
        Dim objType As qbVariableType
        objType = mkType(enuType2)
        If (objType Is Nothing) Then Return (False)
        Dim booResult As Boolean = containedTypeWithState(objType1, objType)
        objType.dispose() : objType = Nothing
        Return (booResult)
    End Function
    ' Unshared version using objects
    Public Function containedTypeWithState(ByVal objType1 As qbVariableType, _
                                           ByVal objType2 As qbVariableType) _
           As Boolean
        If Not checkUsable_("containedTypeWithState") Then Return (False)
        Dim strExplanation As String
        Return (containedTypeWithState(objType1, objType2, strExplanation))
    End Function
    ' Unshared version using objects, with explanation
    Public Function containedTypeWithState(ByVal objType1 As qbVariableType, _
                                           ByVal objType2 As qbVariableType, _
                                           ByRef strExplanation As String) _
           As Boolean
        If Not checkUsable_("containedTypeWithState") Then Return (False)
        Return (containedType_(objType1, _
                              objType2, _
                              USRstate.colTypeOrdering, _
                              USRstate.booContained, _
                              strExplanation))
    End Function
    ' --- Common logic
    Private Shared Function containedType_(ByVal objType1 As qbVariableType, _
                                           ByVal objType2 As qbVariableType, _
                                           ByRef colTypeOrdering As Collection, _
                                           ByRef booContained(,) As Boolean, _
                                           ByRef strExplanation As String) _
           As Boolean
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        strExplanation = "Error has occured"
        If Not objType2.Usable Then
            errorHandler_("An unusable object cannot be tested", _
                          ClassName, _
                          "containedType_", _
                          "Returning False")
            Return (False)
        End If
        If objType1.isUDT OrElse objType2.isUDT Then
            strExplanation = "One or both objects are UDTs"
            Return (False)
        End If
        If (colTypeOrdering Is Nothing) Then
            ' Allocate statics
            Try
                colTypeOrdering = New Collection
                Dim strSplit() As String = Split(TYPEORDERING, " ")
                For intIndex1 = 0 To UBound(strSplit)
                    colTypeOrdering.Add(intIndex1 + 1, strSplit(intIndex1))
                Next intIndex1
                With colTypeOrdering
                    ReDim booContained(.Count, .Count)
                    Dim intIndex3 As Integer
                    strSplit = Split(TYPECONTAINMENT, " ")
                    For intIndex1 = 1 To .Count
                        For intIndex2 = 1 To .Count
                            booContained(intIndex1, intIndex2) = (strSplit(intIndex3) = "Y")
                            intIndex3 += 1
                        Next intIndex2
                    Next intIndex1
                End With
            Catch
                errorHandler_("Cannot create containedType data structures: " & _
                              Err.Number & " " & Err.Description, _
                              ClassName, _
                              "containedType", _
                              "Object isn't usable")
                Return (False)
            End Try
        End If
        If objType1.isUnknown OrElse objType1.isNull _
           OrElse _
           objType2.isUnknown OrElse objType2.isNull Then
            strExplanation = "One or both of the types are Unknown or Null"
            Return (False)
        End If
        Dim strName1 As String = vtPrefixRemove(UCase(objType1.toName))
        Dim strName2 As String = vtPrefixRemove(UCase(objType2.toName))
        If strName2 = "VARIANT" Then
            strExplanation = "Type 2 is a variant and all types can convert to variant"
            Return (True)
        End If
        If strName1 = "VARIANT" Then
            strExplanation = "Type 1 is a variant and type 2 is not a variant"
            Return (False)
        End If
        If strName1 = "ARRAY" Then
            If strName2 <> "ARRAY" Then
                strExplanation = "Type 1 is an array: type 2 is not an array"
                Return (False)
            End If
            Return (containedType_arrayTypes_(objType1, _
                                             objType2, _
                                             colTypeOrdering, _
                                             booContained, _
                                             strExplanation))
        ElseIf strName2 = "ARRAY" Then
            strExplanation = "Type 1 is not an array while type 2 is an array"
            Return (False)
        End If
        intIndex1 = CInt(colTypeOrdering.Item(strName1))
        intIndex2 = CInt(colTypeOrdering.Item(strName2))
        If booContained(intIndex1, intIndex2) Then
            strExplanation = "Any scalar value of type " & strName1 & " " & _
                             "converts without error to type " & strName2
            Return (True)
        End If
        strExplanation = "Some or all scalar values of type " & strName1 & " " & _
                         "do not convert without error to type " & strName2
        Return (False)
    End Function

    ' -----------------------------------------------------------------
    ' Test array type containment on behalf of containedType
    '
    '
    Private Shared Function containedType_arrayTypes_(ByVal objType1 As qbVariableType, _
                                                      ByVal objType2 As qbVariableType, _
                                                      ByVal colTypeOrdering As Collection, _
                                                      ByVal booContained(,) As Boolean, _
                                                      ByRef strExplanation As String) _
            As Boolean
        With objType1
            If .Dimensions <> objType2.Dimensions Then
                strExplanation = "Dimensions of arrays differ"
                Return (False)
            End If
            strExplanation = "Dimensions of arrays are identical"
            Dim booIsomorphic As Boolean
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Dimensions
                If .BoundSize(intIndex1) > objType2.BoundSize(intIndex1) Then
                    strExplanation &= ": " & _
                                      "but, at dimension " & intIndex1 & ", " & _
                                      "type 1's bound size exceeds that of type 2"
                    Return (False)
                End If
                If .LowerBound(intIndex1) <> objType2.LowerBound(intIndex1) _
                   OrElse _
                   .UpperBound(intIndex1) <> objType2.UpperBound(intIndex1) Then
                    booIsomorphic = True
                End If
            Next intIndex1
            If booIsomorphic Then
                strExplanation &= " (note that the arrays are isomorphic with differing " & _
                                  "lower and upper bounds)"
            Else
                strExplanation &= " (note that the arrays have the same bounds)"
            End If
            strExplanation &= ": " & _
                              "and, each type 1 bound size is less than or equal to that of type 2"
            Dim strName1 As String = UCase(.toContainedName)
            Dim strName2 As String = UCase(objType2.toContainedName)
            If strName1 = "VARIANT" Then
                If strName2 = "VARIANT" Then
                    strExplanation &= ": " & _
                                      "and, both type 1 and type 2 contain variant entries"
                    Return (True)
                End If
                strExplanation &= ": " & _
                                  "but, type 1 contains variant entries while type 2 does not"
                Return (True)
            ElseIf strName2 = "VARIANT" Then
                strExplanation &= ": " & _
                                  "and, type 1 contains nonvariant entries while " & _
                                  "type 2 contains variant entries"
            Else
                Dim strInnerExplanation As String
                Dim objContainedType1 As qbVariableType = mkType(string2enuVarType(strName1))
                Dim objContainedType2 As qbVariableType = mkType(string2enuVarType(strName2))
                If containedType_(objContainedType1, _
                                  objContainedType2, _
                                  colTypeOrdering, _
                                  booContained, _
                                  strInnerExplanation) Then
                    strExplanation &= ": " & _
                                      "and, the contained type in array 1 is contained in " & _
                                      "the contained type in array 2: " & _
                                      strInnerExplanation
                    Return (True)
                End If
                strExplanation &= ": " & _
                                    "but, the contained type in array 1 is not contained in " & _
                                    "the contained type in array 2: " & _
                                    strInnerExplanation
                Return (True)
            End If
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return default values
    '
    '
    Public Overloads Function defaultValue() As Object
        If Not checkUsable_("defaultValue") Then Return (Nothing)
        With Me
            If .isArray OrElse .isVariant AndAlso Not .Abstract Then Return (.defaultValue(.VarType))
        End With
        Return defaultValue(USRstate.enuVariableType, Me.Name)
    End Function
    Public Overloads Shared Function defaultValue(ByVal enuType As ENUvarType) _
           As Object
        Return defaultValue(enuType, ClassName)
    End Function
    Private Overloads Shared Function defaultValue(ByVal enuType As ENUvarType, _
                                                   ByVal strObjectName As String) _
            As Object
        Select Case enuType
            Case ENUvarType.vtBoolean : Return (CBool(DEFAULT_BOOLEAN))
            Case ENUvarType.vtByte : Return (CByte(DEFAULT_BYTE))
            Case ENUvarType.vtInteger : Return (CShort(DEFAULT_INTEGER))
            Case ENUvarType.vtLong : Return (CInt(DEFAULT_LONG))
            Case ENUvarType.vtSingle : Return (CSng(DEFAULT_SINGLE))
            Case ENUvarType.vtDouble : Return (CDbl(DEFAULT_DOUBLE))
            Case ENUvarType.vtString : Return (DEFAULT_STRING)
            Case ENUvarType.vtVariant : Return (DEFAULT_VARIANT)
            Case ENUvarType.vtUDT : Return (DEFAULT_UDT)
            Case ENUvarType.vtNull : Return (Nothing)
            Case ENUvarType.vtUnknown
                Return (Nothing)
            Case ENUvarType.vtArray
                errorHandler_("Cannot obtain a default value for an Array using this syntax", _
                              ClassName, _
                              "defaultValue", _
                              "Returning Nothing")
                Return (Nothing)
            Case Else
                errorHandler_("Internal programming error: unexpected varType is " & _
                              _OBJutilities.enquote(enuType.ToString), _
                              ClassName, _
                              "defaultValue", _
                              "Returning Nothing")
                Return (Nothing)
        End Select
    End Function


    ' ----------------------------------------------------------------------
    ' Return the number of dimensions
    '
    '
    Public ReadOnly Property Dimensions() As Integer
        Get
            If Not checkUsable_("Dimensions get") Then Return (0)
            With USRstate
                If (.colBounds Is Nothing) Then Return (0)
                Return (.colBounds.Count)
            End With
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Dispose of the object
    '
    '
    ' --- Conduct an inspection
    Public Sub dispose() Implements IDisposable.Dispose
        dispose(False)
    End Sub
    ' --- Conduct an inspection
    Public Function disposeInspect() As Boolean
        Return (dispose(True))
    End Function
    ' --- Exposes option to dispose
    Private Function dispose(ByVal booInspect As Boolean) As Boolean
        If Not checkUsable_("dispose") Then Return (False)
        If booInspect Then inspection_(False)
        With USRstate
            If Not (.colBounds Is Nothing) Then
                With .colBounds
                    Dim intIndex1 As Integer
                    Dim colHandle As Collection
                    For intIndex1 = .Count To 1 Step -1
                        colHandle = CType(.Item(intIndex1), Collection)
                        colHandle = Nothing
                        .Remove(intIndex1)
                    Next intIndex1
                End With
                .colBounds = Nothing
            End If
            If Not (.objVarType Is Nothing) Then
                If Me.isUDT Then
                    If Not clearUDTmembers_() Then Return (False)
                Else
                    If Not CType(.objVarType, qbVariableType).disposeInspect Then Return (False)
                    .objVarType = Nothing
                End If
            End If
            If Not (OBJcollectionUtilities Is Nothing) Then
                If Not OBJcollectionUtilities.dispose Then Return (False)
                OBJcollectionUtilities = Nothing
            End If
            Me.mkUnusable()
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Convert the type from a string
    '
    '
    Public Function fromString(ByVal strFromString As String) _
           As Boolean
        If Not checkUsable_("fromString") Then Return (False)
        If Trim(strFromString) = "" Then
            errorHandler_("fromString is null", _
                          "fromString", _
                          "Returning False: making no change to object")
            Return (False)
        End If
        If Not reset_() Then Return (False)
        If cacheCheck_(strFromString) Then Return True
        Dim strFromStringWork As String = strFromString
        Select Case UCase(vtPrefixAdd(Trim(strFromString)))
            Case "VTVARIANT" : strFromStringWork &= ",Null"
            Case "VTARRAY" : strFromStringWork &= ",Variant,0,0"
        End Select
        Dim objScanner As qbScanner.qbScanner
        Try
            objScanner = New qbScanner.qbScanner
        Catch
            errorHandler_("Cannot make a scanner: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString", _
                          "Marking object unusable and returning False")
            Me.mkUnusable() : Return (False)
        End Try
        With objScanner
            .SourceCode = strFromString
            If Not .scan Then
                errorHandler_("Cannot scan source code", _
                              "fromString", _
                              "Returning False")
                Me.mkUnusable() : Return (False)
            End If
        End With
        Dim intIndex1 As Integer
        Dim strError As String
        If Not fromString_parse_(strFromStringWork, objScanner, strError, intIndex1) Then
            errorHandler_("Cannot parse fromString " & _
                          _OBJutilities.enquote(strFromString) & vbNewLine & vbNewLine & _
                          "Error at or near " & intIndex1 & ": " & _
                          strError, _
                          "fromString", _
                          "Resetting object to unknown status")
            reset_()
            Return (False)
        End If
        cacheUpdate_(strFromString)
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Parse the fromString on behalf of the fromString method
    '
    '
    Private Function fromString_parse_(ByVal strFromString As String, _
                                       ByVal objScanner As qbScanner.qbScanner, _
                                       ByRef strError As String, _
                                       ByRef intFinalIndex As Integer) _
            As Boolean
        intFinalIndex = 1
        Return (fromString_parse__typeSpecification_(intFinalIndex, objScanner, strError) _
               AndAlso _
               intFinalIndex > objScanner.TokenCount)
    End Function

    ' ----------------------------------------------------------------------
    ' abstractVariantType := [VT] VARIANT
    '
    '
    Private Function fromString_parse__abstractVariantType_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef strError As String) _
            As Boolean
        Return (fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             "VARIANT", _
                                             booVTprefix:=True))
    End Function

    ' ----------------------------------------------------------------------
    ' arrayType := [VT] ARRAY,arrType,boundList
    '
    '
    Private Function fromString_parse__arrayType_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objContainer As qbVariableType, _
             ByRef objContained As qbVariableType, _
             ByRef strError As String) _
            As Boolean
        Dim intIndex1 As Integer = intIndex
        Try
            objContained = New qbVariableType
        Catch
            errorHandler_("Cannot create contained variable type for array: " & _
                          Err.Number & " " & Err.Description, _
                          "", _
                          "Returning False")
        End Try
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             "ARRAY", _
                                             booVTprefix:=True) _
           OrElse _
           Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) _
           OrElse _
           Not fromString_parse__arrType_(intIndex, _
                                          objScanner, _
                                          objContainer, _
                                          objContained, _
                                          strError) _
           OrElse _
           Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) _
           OrElse _
           Not fromString_parse__boundList_(intIndex, _
                                            objScanner, _
                                            objContainer, _
                                            strError) Then
            intIndex = intIndex1
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Array declaration is not valid")
            Return (False)
        End If
        objContainer.VariableTypeEnum = ENUvarType.vtArray
        objContainer.ContainedVariableType = objContained
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' arrType := simpleType | abstractVariantType | parUDT
    '
    '
    Private Function fromString_parse__arrType_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objContainer As qbVariableType, _
             ByRef objContained As qbVariableType, _
             ByRef strError As String) _
            As Boolean
        If fromString_parse__simpleType_(intIndex, _
                                         objScanner, _
                                         objContained, _
                                         strError) Then
            Return (True)
        End If
        If fromString_parse__abstractVariantType_(intIndex, objScanner, strError) Then
            objContained.VariableTypeEnum = ENUvarType.vtVariant
            Return (True)
        End If
        Dim colContained As Collection
        If fromString_parse__parUDT_(intIndex, objScanner, objContained, strError) Then
            Return (True)
        End If
        _OBJutilities.append(strError, _
                             vbNewLine, _
                             "Array declaration contains an invalid type")
        Return (False)
    End Function

    ' ----------------------------------------------------------------------
    ' baseType := simpleType | variantType | arrayType
    '
    '
    Private Function fromString_parse__baseType_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objContainer As qbVariableType, _
             ByRef objContained As qbVariableType, _
             ByRef strError As String) _
            As Boolean
        If fromString_parse__simpleType_(intIndex, objScanner, objContainer, strError) Then
            Return (True)
        End If
        Try
            objContained = New qbVariableType
        Catch
            errorHandler_("Cannot create contained type: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString_parse__typeSpecification_", _
                          "Returning False")
        End Try
        If fromString_parse__variantType_(intIndex, _
                                          objScanner, _
                                          objContained, _
                                          strError) Then
            objContainer.VariableTypeEnum = ENUvarType.vtVariant
            Return (True)
        End If
        If fromString_parse__arrayType_(intIndex, _
                                        objScanner, _
                                        objContainer, _
                                        objContained, _
                                        strError) Then
            objContainer.VariableTypeEnum = ENUvarType.vtArray
            Return (True)
        End If
        _OBJutilities.append(strError, _
                             vbNewLine, _
                             "A simple type, variant or an array was expected and not found")
        Return (False)
    End Function

    ' ----------------------------------------------------------------------
    ' boundList := boundListEntry | boundListEntry COMMA boundList
    '
    '
    Private Function fromString_parse__boundList_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objQBvariableType As qbVariableType, _
             ByRef strError As String) _
            As Boolean
        Dim intLBound As Integer
        Dim intUBound As Integer
        If Not fromString_parse__boundListEntry_(intIndex, _
                                                 objScanner, _
                                                 intLBound, intUBound, _
                                                 strError) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "boundList doesn't start with a bound list entry")
            Return (False)
        End If
        With objQBvariableType
            If Not .createBoundList_ Then Return (False)
            If Not .extendBoundList_(intLBound, intUBound) Then Return (False)
        End With
        Do
            If Not fromString_parse__checkToken_(intIndex, _
                                                 objScanner, _
                                                 qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) Then
                Exit Do
            End If
            If Not fromString_parse__boundListEntry_(intIndex, _
                                                     objScanner, _
                                                     intLBound, intUBound, _
                                                     strError) Then
                errorHandler_("Syntax error in variable type " & _
                              _OBJutilities.enquote(objScanner.SourceCode) & ": " & _
                              "comma at position " & _
                              objScanner.tokenStartIndex(intIndex - 1) & " " & _
                              "is not followed by a bound list entry", _
                              "fromString_parse__boundList_", "")
                Exit Do
            End If
            If Not objQBvariableType.extendBoundList_(intLBound, intUBound) Then
                Return (False)
            End If
        Loop
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' boundListEntry := BOUNDINTEGER,BOUNDINTEGER
    '
    '
    Private Function fromString_parse__boundListEntry_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef intLBound As Integer, _
             ByRef intUBound As Integer, _
             ByRef strError As String) _
            As Boolean
        Dim intLBoundWork As Integer
        Dim intSignum As Integer = fromString_parse__boundListEntrySignum_ _
                                   (intIndex, objScanner)
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Lower bound is not a number")
            Return (False)
        End If
        With objScanner
            intLBoundWork = CInt(.sourceMid(intIndex - 1)) * intSignum
        End With
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Lower bound is not followed by a comma")
            Return (False)
        End If
        intSignum = fromString_parse__boundListEntrySignum_(intIndex, objScanner)
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Upper bound is not a number")
            Return (False)
        End If
        intLBound = intLBoundWork
        With objScanner
            intUBound = CInt(.sourceMid(intIndex - 1)) * intSignum
        End With
        If intLBound > intUBound Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Lower and upper bounds are in the wrong order")
            Return (False)
        End If
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Check for the sign of the bound list entry
    '
    '
    Private Function fromString_parse__boundListEntrySignum_ _
                     (ByRef intIndex As Integer, ByVal objScanner As qbScanner.qbScanner) _
            As Integer
        If fromString_parse__checkToken_(intIndex, objScanner, "+") Then
            Return (1)
        ElseIf fromString_parse__checkToken_(intIndex, objScanner, "-") Then
            Return (-1)
        Else
            Return (1)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Check for a specific token or token type
    '
    '
    Private Overloads Function fromString_parse__checkToken_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByVal enuExpected As qbTokenType.qbTokenType.ENUtokenType) As Boolean
        Return (objScanner.checkToken(intIndex, enuExpected))
    End Function
    Private Overloads Function fromString_parse__checkToken_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByVal strExpected As String, _
             Optional ByVal booVTprefix As Boolean = False) As Boolean
        With objScanner
            Return (.checkToken(intIndex, strExpected) _
                   OrElse _
                   booVTprefix AndAlso .checkToken(intIndex, vtPrefixAdd(strExpected)))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' parMemberType := LEFTPAR MEMBERNAME,baseType RIGHTPAR
    '
    '
    Private Function fromString_parse__parBaseType_(ByRef intIndex As Integer, _
                                                    ByVal objScanner As qbScanner.qbScanner, _
                                                    ByRef strMemberName As String, _
                                                    ByRef objBaseType As qbVariableType, _
                                                    ByRef strError As String) _
            As Boolean
        If Not fromString_parse__checkToken_(intIndex, objScanner, "(") Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Expected left parenthesis not found")
            Return (False)
        End If
        Dim intEndIndex As Integer = objScanner.findRightParenthesis(intIndex)
        If intEndIndex > objScanner.TokenCount Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT's base type does not contain a balanced right parenthesis")
            Return (False)
        End If
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier) _
           OrElse _
           Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT's base type does not start with a member name and a comma")
            Return (False)
        End If
        Dim strFromString As String
        With objScanner
            strMemberName = .sourceMid(intIndex - 2)
            strFromString = .sourceMid(intIndex, intEndIndex - intIndex)
            Dim intLength As Integer = Len(strFromString)
            If Mid(strFromString, 1, 1) = "(" AndAlso Mid(strFromString, intLength) = ")" Then
                strFromString = Mid(strFromString, 2, intLength - 2)
            End If
        End With
        Try
            objBaseType = New qbVariableType(strFromString)
        Catch
            errorHandler_("Cannot create the UDT base type from " & _
                          _OBJutilities.enquote(strFromString) & ": " & _
                          Err.Number & " " & Err.Description, _
                          "", _
                          "Returning False")
            Return (False)
        End Try
        intIndex = intEndIndex + 1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' parUDT := LEFTPARENTHESIS udt RIGHTPARENTHESIS
    '
    '
    Private Function fromString_parse__parUDT_(ByRef intIndex As Integer, _
                                               ByVal objScanner As qbScanner.qbScanner, _
                                               ByRef objContainer As qbVariableType, _
                                               ByRef strError As String) _
            As Boolean
        Dim intIndex1 As Integer = intIndex
        If Not fromString_parse__checkToken_(intIndex, objScanner, "(") Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Missing left parenthesis")
            Return (False)
        End If
        Dim colContained As Collection
        If Not fromString_parse__UDT_(intIndex, _
                                      objScanner, _
                                      objContainer, _
                                      colContained, _
                                      strError) Then
            intIndex = intIndex1
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT definition not found")
            Return (False)
        End If
        If Not fromString_parse__checkToken_(intIndex, objScanner, ")") Then
            intIndex = intIndex1
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Missing right parenthesis")
            Return (False)
        End If
        With objContainer
            .ContainedVariableType = colContained
            .VariableTypeEnum = ENUvarType.vtUDT
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' simpleType := [VT] typeName
    ' simpleType := BOOLEAN|BYTE|INTEGER|LONG|SINGLE|DOUBLE|STRING|
    '               UNKNOWN|NULL
    '
    '
    Private Function fromString_parse__simpleType_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objQBvariableType As qbVariableType, _
             ByRef strError As String) _
            As Boolean
        With objQBvariableType
            Select Case UCase(vtPrefixAdd(Mid(objScanner.SourceCode, _
                                               objScanner.tokenStartIndex(intIndex), _
                                               objScanner.tokenLength(intIndex))))
                Case "VTBOOLEAN" : .VariableTypeEnum = ENUvarType.vtBoolean
                Case "VTBYTE" : .VariableTypeEnum = ENUvarType.vtByte
                Case "VTINTEGER" : .VariableTypeEnum = ENUvarType.vtInteger
                Case "VTLONG" : .VariableTypeEnum = ENUvarType.vtLong
                Case "VTSINGLE" : .VariableTypeEnum = ENUvarType.vtSingle
                Case "VTDOUBLE" : .VariableTypeEnum = ENUvarType.vtDouble
                Case "VTSTRING" : .VariableTypeEnum = ENUvarType.vtString
                Case "VTUNKNOWN" : .VariableTypeEnum = ENUvarType.vtUnknown
                Case "VTNULL" : .VariableTypeEnum = ENUvarType.vtNull
                Case Else
                    _OBJutilities.append(strError, _
                                         vbNewLine, _
                                         "Invalid simple type")
                    Return (False)
            End Select
            intIndex += 1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' typeList := parMemberType [ COMMA typeList ]
    '
    '
    Private Function fromString_parse__typeList_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef colContained As Collection, _
             ByRef strError As String) _
            As Boolean
        Dim objNext As qbVariableType
        Dim strMemberName As String
        If Not fromString_parse__parBaseType_(intIndex, _
                                              objScanner, _
                                              strMemberName, _
                                              objNext, _
                                              strError) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT type list doesn't start with a parenthesized member and its type")
            Return (False)
        End If
        Try
            Dim colEntry As Collection
            colEntry = New Collection
            With colEntry
                .Add(colContained.Count + 1)
                .Add(strMemberName)
                .Add(objNext)
            End With
            colContained.Add(colEntry, strMemberName)
        Catch
            errorHandler_("Cannot attach UDT member: " & Err.Number & " " & Err.Description, _
                          "fromString_parse__typeList_", _
                          "Marking object unusable, and returning False")
            Me.mkUnusable() : Return (False)
        End Try
        If fromString_parse__checkToken_(intIndex, _
                                         objScanner, _
                                         qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) Then
            If Not fromString_parse__typeList_(intIndex, objScanner, colContained, strError) Then
                _OBJutilities.append(strError, _
                                     vbNewLine, _
                                     "Comma is not followed by a type list entry")
                Return (False)
            End If
        End If
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' typeSpecification := baseType | udt
    '
    '
    Private Function fromString_parse__typeSpecification_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef strError As String) _
            As Boolean
        Dim objContained As qbVariableType
        If fromString_parse__baseType_(intIndex, _
                                       objScanner, _
                                       Me, _
                                       objContained, _
                                       strError) Then
            USRstate.objVarType = objContained
            Return (True)
        End If
        Dim colContained As Collection
        If fromString_parse__UDT_(intIndex, _
                                  objScanner, _
                                  Me, _
                                  colContained, _
                                  strError) Then
            USRstate.objVarType = colContained
            USRstate.enuVariableType = ENUvarType.vtUDT
            Return (True)
        End If
        _OBJutilities.append(strError, _
                             vbNewLine, _
                             "Invalid type specification")
        Return (False)
    End Function

    ' ----------------------------------------------------------------------
    ' udt := [VT] UDT,typeList
    '
    '
    Private Function fromString_parse__UDT_ _
            (ByRef intIndex As Integer, _
             ByVal objScanner As qbScanner.qbScanner, _
             ByRef objContainer As qbVariableType, _
             ByRef colContained As Collection, _
             ByVal strError As String) _
            As Boolean
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             "UDT", _
                                             booVTprefix:=True) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT does not start with keyword UDT")
            Return (False)
        End If
        If Not fromString_parse__checkToken_(intIndex, _
                                             objScanner, _
                                             qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "UDT in user data type declaration is not followed by a comma")
            Return (False)
        End If
        Try
            colContained = New Collection
        Catch
            errorHandler_("Cannot create collection for UDT members: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString_parse__UDT_", _
                          "Marking object unusable: returning True")
        End Try
        Return (fromString_parse__typeList_(intIndex, objScanner, colContained, strError))
    End Function

    ' -----------------------------------------------------------------
    ' variantType := abstractVariantType COMMA varType
    '
    '
    Private Function fromString_parse__variantType_(ByRef intIndex As Integer, _
                                                    ByVal objScanner As qbScanner.qbScanner, _
                                                    ByRef objQBvariableType As qbVariableType, _
                                                    ByRef strError As String) _
            As Boolean
        If Not fromString_parse__abstractVariantType_(intIndex, objScanner, strError) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Variant type does not start with the keyword Variant")
            Return (False)
        End If
        If intIndex > objScanner.TokenCount Then
            objQBvariableType.changeVariableType(ENUvarType.vtNull) : Return (True)
        End If
        If Not (fromString_parse__checkToken_(intIndex, _
                                              objScanner, _
                                              qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma) _
                AndAlso _
                fromString_parse__varType_(intIndex, _
                                           objScanner, _
                                           objQBvariableType, _
                                           strError)) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "Concrete variant syntax is not 'Variant,type'")
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' varType := simpleType|(arrayType)
    '
    '
    Private Function fromString_parse__varType_(ByRef intIndex As Integer, _
                                                ByVal objScanner As qbScanner.qbScanner, _
                                                ByRef objQBvariableType As qbVariableType, _
                                                ByRef strError As String) _
            As Boolean
        If fromString_parse__simpleType_(intIndex, _
                                         objScanner, _
                                         objQBvariableType, _
                                         strError) Then Return (True)
        If Not (fromString_parse__checkToken_(intIndex, objScanner, "(") _
                AndAlso _
                fromString_parse__arrayType_(intIndex, _
                                             objScanner, _
                                             objQBvariableType, _
                                             CType(objQBvariableType.ContainedVariableType, _
                                                   qbVariableType), _
                                             strError) _
                AndAlso _
                fromString_parse__checkToken_(intIndex, objScanner, ")")) Then
            _OBJutilities.append(strError, _
                                 vbNewLine, _
                                 "The contained variant type is not a simple type or an array")
            Return (False)
        End If
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Rebuild fromString input from the scanner: remove VT prefixes
    '
    '
    Private Function fromString_rebuild_(ByVal objScanner As qbScanner.qbScanner) _
            As String
        With objScanner
            Return (.sourceMid(1, .tokenEndIndex(1) - .tokenStartIndex(.TokenCount) + 1))
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Extract the typename
    '
    '
    Public Shared Function fromString2TypeName(ByVal strFromstring As String) _
           As String
        If _OBJutilities.words(strFromstring) = 1 Then Return strFromstring
        Return _OBJutilities.listItem(strFromstring, 2)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert the type enumerator to its Hungarian prefix
    '
    '
    Public Shared Function hungarianPrefix(ByVal enuType As ENUvarType) As String
        Select Case vtPrefixRemove(UCase(enuType.ToString))
            Case "BOOLEAN" : Return "boo"
            Case "BYTE" : Return "byt"
            Case "INTEGER" : Return "int"
            Case "LONG" : Return "lng"
            Case "SINGLE" : Return "sgl"
            Case "DOUBLE" : Return "dbl"
            Case "STRING" : Return "str"
            Case "VARIANT" : Return "var"
            Case "ARRAY" : Return "arr"
            Case "UDT" : Return "typ"
            Case "UNKNOWN" : Return "unk"
            Case "NULL" : Return "nul"
            Case Else
                errorHandler_("Unexpected type " & enuType.ToString, _
                              ClassName, "hungarianPrefix", _
                              "Returning null string")
                Return ("")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Return contained type
    '
    '
    ' --- Not indexed: get Variant or array type
    Public Overloads Function innerType() As qbVariableType
        If Not checkUsable_("innerType") Then Return (Nothing)
        If Me.isUDT Then
            errorHandler_("Syntax is not valid for UDT: requires an index", _
                          "innerType", _
                          "Returning Nothing")
            Return (Nothing)
        End If
        Return (innerType_(0))
    End Function
    ' --- Indexed: get UDT member
    Public Overloads Function innerType(ByVal objIndex As Object) As qbVariableType
        If Not Me.isUDT Then
            errorHandler_("Syntax is not valid for non-UDT: index cannot be specified", _
                          "innerType", _
                          "Returning Nothing")
            Return (Nothing)
        End If
        Dim intIndex1 As Integer = udtIndex2Integer_(objIndex, _
                                                     CType(USRstate.objVarType, _
                                                           Collection).Count)
        If intIndex1 = 0 Then Return (Nothing)
        Return (innerType_(intIndex1))
    End Function
    ' --- Common logic
    Private Function innerType_(ByVal intIndex As Integer) As qbVariableType
        If Not checkUsable_("innerType") Then Return (Nothing)
        With USRstate
            Dim objHandle As qbVariableType
            If intIndex <> 0 Then
                Dim colHandle As Collection = CType(.objVarType, Collection)
                With colHandle
                    objHandle = CType(CType(.Item(intIndex), Collection).Item(3), qbVariableType)
                End With
            Else
                objHandle = CType(.objVarType, qbVariableType)
            End If
            Return (objHandle)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Inspection
    '
    '
    Public Overridable Function inspect(ByRef strReport As String, _
                                        Optional ByVal booBasic As Boolean = False) As Boolean
        Dim booInspection As Boolean = True
        strReport = "Inspection of " & _
                    _OBJutilities.enquote(Me.Name) & " " & _
                    "(" & Me.ToString & ") " & _
                    "at " & Now & _
                    vbNewLine & vbNewLine
        If Not _OBJutilities.inspectionAppend(strReport, _
                                             INSPECTION_USABLE, _
                                             Me.Usable, _
                                             booInspection) Then
            Return (False)
        End If
        With USRstate
            Dim objVarType As qbVariableType
            If Not Me.isUDT Then
                objVarType = CType(.objVarType, qbVariableType)
            End If
            _OBJutilities.inspectionAppend(strReport, _
                                            INSPECTION_TYPECOMPATIBLE, _
                                            (.objVarType Is Nothing) _
                                            AndAlso _
                                            isSimpleType_(.enuVariableType) _
                                            OrElse _
                                            .enuVariableType = ENUvarType.vtArray _
                                            AndAlso _
                                            (isScalarType(objVarType.VariableTypeEnum) _
                                             OrElse _
                                             objVarType.VariableTypeEnum = ENUvarType.vtVariant _
                                             OrElse _
                                             objVarType.VariableTypeEnum = ENUvarType.vtUDT) _
                                            AndAlso _
                                            Me.Dimensions > 0 _
                                            OrElse _
                                            .enuVariableType = ENUvarType.vtVariant _
                                            AndAlso _
                                            (Me.Abstract _
                                             OrElse _
                                             (isScalarType(objVarType.VariableTypeEnum) _
                                              OrElse _
                                              objVarType.VariableTypeEnum = ENUvarType.vtArray _
                                              OrElse _
                                              objVarType.VariableTypeEnum = ENUvarType.vtUnknown _
                                              OrElse _
                                              objVarType.VariableTypeEnum = ENUvarType.vtNull)) _
                                             OrElse _
                                             .enuVariableType = ENUvarType.vtUDT _
                                             AndAlso _
                                             (TypeOf .objVarType Is Collection), _
                                            booInspection)
            If Not booBasic AndAlso Not Me.Abstract AndAlso Not Me.isUDT Then
                _OBJutilities.inspectionAppend(strReport, _
                                              INSPECTION_CLONEOK, _
                                              inspect_cloneOK_, _
                                              booInspection)
                _OBJutilities.inspectionAppend(strReport, _
                                              INSPECTION_TOFROMSTRING, _
                                              checkToFromString_(Me.ToString), _
                                              booInspection)
            End If
            If Not (.objVarType Is Nothing) Then
                Dim strSubreport As String
                If Me.isUDT Then
                    If _OBJutilities.inspectionAppend(strReport, _
                                                      INSPECTION_TYPECOMPATIBLE, _
                                                      (TypeOf .objVarType Is Collection), _
                                                      booInspection, _
                                                      "Since the container is a UDT, " & _
                                                      "the contained type should be a " & _
                                                      "collection of members") Then
                        Dim strExplanation As String
                        Dim booOK As Boolean = inspect_udtCollection_(CType(.objVarType, Collection), _
                                                                      strExplanation)
                        _OBJutilities.inspectionAppend(strReport, _
                                                       INSPECTION_CONTAINEDOK, _
                                                       booOK, _
                                                       booInspection)
                    End If
                Else
                    Dim objContainedVarType As qbVariableType = CType(.objVarType, qbVariableType)
                    _OBJutilities.inspectionAppend(strReport, _
                                                   INSPECTION_CONTAINEDOK, _
                                                   objContainedVarType.inspect(strSubreport), _
                                                   booInspection)
                    strReport &= vbNewLine & vbNewLine & _
                                _OBJutilities.string2Box(strSubreport, _
                                                         "Inspection of contained data type object " & _
                                                         objContainedVarType.Name)
                End If
            End If
            Return (booInspection)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Test OK clone on behalf of inspect
    '
    '
    Private Function inspect_cloneOK_() As Boolean
        Dim objClone As qbVariableType
        Try
            objClone = Me.clone
            objClone.fromString(Me.ToString)
        Catch
            _OBJutilities.errorHandler("Cannot make clone: " & _
                                      Err.Number & " " & Err.Description, _
                                      "", _
                                      "Returning False and marking object " & _
                                      "not usable")
            Me.mkUnusable() : Return (False)
        End Try
        Return (Me.compareTo(objClone))
    End Function

    ' ----------------------------------------------------------------------
    ' Inspect the UDT collection
    '
    '
    Public Function inspect_udtCollection_(ByVal colUDT As Collection, _
                                           ByRef strExplanation As String) _
           As Boolean
        Dim booOK As Boolean
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strReport As String
        With colUDT
            For intIndex1 = 1 To .Count
                If Not (TypeOf .Item(intIndex1) Is Collection) Then
                    strExplanation = "Item " & intIndex1 & " is not a collection"
                    Return (False)
                End If
                Dim colItem As Collection = CType(.Item(intIndex1), Collection)
                With colItem
                    If .Count <> 3 Then
                        strExplanation = "Item " & intIndex1 & " is not a 3-item collection"
                        Return (False)
                    End If
                    booOK = False
                    If (TypeOf .Item(1) Is System.Int32) Then
                        intIndex2 = CInt(.Item(1))
                        booOK = (intIndex2 = intIndex1)
                    End If
                    If Not booOK Then
                        strExplanation = "Item " & _
                                         intIndex1 & " " & _
                                         "doesn't contain the member index in item(1)"
                        Return (False)
                    End If
                    booOK = False
                    If (TypeOf .Item(2) Is System.String) AndAlso Not IsNumeric(.Item(2)) Then
                        Try
                            Dim colHandle As Collection = CType(colUDT.Item(CStr(.Item(2))), _
                                                                Collection)
                            booOK = True
                        Catch : End Try
                    End If
                    If Not booOK Then
                        strExplanation = "Item " & _
                                         intIndex1 & " " & _
                                         "doesn't contain a member name in item(2)"
                        Return (False)
                    End If
                    If Not (TypeOf .Item(3) Is qbVariableType) Then
                        strExplanation = "Item " & _
                                         intIndex1 & " " & _
                                         "doesn't contain a qbVariableType object in item(2)"
                        Return (False)
                    End If
                    If Not CType(.Item(3), qbVariableType).inspect(strReport) Then
                        strExplanation = "qbVariableType object in item " & _
                                         intIndex1 & " " & _
                                         "fails its inspection: " & _
                                         vbNewLine & vbNewLine & _
                                         strReport
                        Return (False)
                    End If
                End With
            Next intIndex1
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable is array
    '
    '
    Public Function isArray() As Boolean
        If Not checkUsable_("isArray") Then Return (False)
        Return (Me.Dimensions > 0)
    End Function

    ' -----------------------------------------------------------------
    ' Indicate use of default values
    '
    '
    ' --- Stateful version specifies .Net value only
    Public Overloads Function isDefaultValue(ByVal objNetValue As Object) As Boolean
        If Not checkUsable_("defaultValue") Then Return (False)
        Return isDefaultValue(objNetValue, _
                              CType(IIf(Me.isArray, Me.VarType, Me.VariableType), _
                                    ENUvarType), _
                              Me.Name)
    End Function
    ' --- Stateless version specifies type     
    Public Overloads Shared Function isDefaultValue(ByVal objNetValue As Object, _
                                                    ByVal enuType As ENUvarType) _
           As Boolean
        Return isDefaultValue(objNetValue, _
                              enuType, _
                              ClassName)
    End Function
    ' --- Common functionality  
    Private Overloads Shared Function isDefaultValue(ByVal objNetValue As Object, _
                                                     ByVal enuType As ENUvarType, _
                                                     ByVal strObjectName As String) _
            As Boolean
        Try
            Dim objDefault As Object = defaultValue(enuType)
            Select Case enuType
                Case ENUvarType.vtBoolean
                    Return (CBool(objNetValue) = CBool(objDefault))
                Case ENUvarType.vtByte
                    Return (CByte(objNetValue) = CByte(objDefault))
                Case ENUvarType.vtInteger
                    Return (CShort(objNetValue) = CShort(objDefault))
                Case ENUvarType.vtLong
                    Return (CLng(objNetValue) = CLng(objDefault))
                Case ENUvarType.vtSingle
                    Return (CSng(objNetValue) = CSng(objDefault))
                Case ENUvarType.vtDouble
                    Return (CDbl(objNetValue) = CDbl(objDefault))
                Case ENUvarType.vtString
                    Return (CStr(objNetValue) = CStr(objDefault))
                Case ENUvarType.vtVariant
                    Return (objNetValue Is Nothing)
                Case ENUvarType.vtNull
                    Return (objDefault Is Nothing)
                Case Else
                    errorHandler_("Internal programming error: unexpected varType is " & _
                                  _OBJutilities.enquote(enuType.ToString), _
                                  strObjectName, _
                                  "defaultValue", _
                                  "Returning Nothing")
                    Return (False)
            End Select
        Catch
            Return (False)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable is Null
    '
    '
    Public Function isNull() As Boolean
        If Not checkUsable_("isNull") Then Return (False)
        Return (USRstate.enuVariableType = ENUvarType.vtNull)
    End Function

    ' ----------------------------------------------------------------------
    ' Determine type isomorphism
    '
    '
    ' --- No explanation
    Public Overloads Function isomorphicType(ByVal objType2 As qbVariableType) _
           As Boolean
        Dim strToss As String
        Return (isomorphicType(objType2, strToss))
    End Function
    ' --- Explanation provided
    Public Overloads Function isomorphicType(ByVal objType2 As qbVariableType, _
                                             ByRef strExplanation As String) _
           As Boolean
        If Not checkUsable_("isomorphicType") Then
            strExplanation = "qbVariable object is not usable"
            Return (False)
        End If
        Dim strExplanation2 As String
        Dim booIsomorphic As Boolean = _
            Me.containedTypeWithState(Me, objType2, strExplanation) _
            AndAlso _
            Me.containedTypeWithState(objType2, Me, strExplanation2)
        strExplanation &= ": " & strExplanation2
        Return (booIsomorphic)
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable is scalar
    '
    '
    Public Function isScalar() As Boolean
        If Not checkUsable_("isArray") Then Return (False)
        Return (Me.Dimensions = 0 _
                AndAlso _
                Not Me.isVariant _
                AndAlso _
                Not _
                Me.isNull _
                AndAlso _
                Not Me.isUnknown _
                AndAlso _
                Not Me.isUDT)
    End Function

    ' -----------------------------------------------------------------------
    ' Tell caller if type is scalar
    '
    '
    Public Overloads Shared Function isScalarType(ByVal enuType As ENUvarType) As Boolean
        Return (isScalarType(enuType.ToString))
    End Function
    Public Overloads Shared Function isScalarType(ByVal strType As String) As Boolean
        Select Case UCase(Trim(vtPrefixAdd(strType)))
            Case "VTBOOLEAN"
            Case "VTBYTE"
            Case "VTINTEGER"
            Case "VTLONG"
            Case "VTSINGLE"
            Case "VTDOUBLE"
            Case "VTSTRING"
            Case Else : Return (False)
        End Select
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable type is UDT
    '
    '
    Public Function isUDT() As Boolean
        If Not checkUsable_("isUDT") Then Return (False)
        With USRstate
            If (.objVarType Is Nothing) Then Return (False)
            If (TypeOf .objVarType Is qbVariableType) Then Return (False)
            If (TypeOf .objVarType Is Collection) AndAlso Me.Dimensions = 0 Then Return (True)
            errorHandler_("State contains an unexpected objVarType", _
                          "isUDT", _
                          "Marking object unusable and returning False")
            Me.mkUnusable() : Return (False)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable type is Unknown
    '
    '
    Public Function isUnknown() As Boolean
        If Not checkUsable_("isUnknown") Then Return (False)
        Return (USRstate.enuVariableType = ENUvarType.vtUnknown)
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller if variable is Variant
    '
    '
    Public Function isVariant() As Boolean
        If Not checkUsable_("isVariant") Then Return (False)
        Return (USRstate.enuVariableType = ENUvarType.vtVariant)
    End Function

    ' ----------------------------------------------------------------------
    ' Return the lowerBound at the indicated dimension
    '
    '
    Public Property LowerBound(ByVal intDimension As Integer) As Integer
        Get
            If Not checkUsable_("LowerBound get") Then Return (0)
            If Me.VariableType <> ENUvarType.vtArray Then
                errorHandler_("Lower bound requested for nonarray", _
                              "LowerBound get", _
                              "Returning 0")
                Return (0)
            End If
            If intDimension < 1 OrElse intDimension > Me.Dimensions Then
                errorHandler_("Dimension requested for lower bound is less than 1 " & _
                              "or greater than dimensions " & Me.Dimensions, _
                              "LowerBound get", _
                              "Returning 0")
                Return (0)
            End If
            Return (CInt(CType(USRstate.colBounds.Item(intDimension), Collection).Item(1)))
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("LowerBound set") Then Return
            If Me.VariableType <> ENUvarType.vtArray Then
                errorHandler_("Lower bound requested for nonarray", _
                              "LowerBound set", _
                              "No change made")
                Return
            End If
            If intDimension < 1 OrElse intDimension > Me.Dimensions Then
                errorHandler_("Dimension requested for lower bound is less than 1 " & _
                              "or greater than dimensions " & Me.Dimensions, _
                              "LowerBound set", _
                              "No change made")
                Return
            End If
            Dim intUpperBound As Integer = Me.UpperBound(intDimension)
            If intNewValue > intUpperBound Then
                errorHandler_("Proposed lower bound " & intNewValue & " " & _
                              "is greater than actual upper bound" & _
                              intUpperBound, _
                              "LowerBound set", _
                              "No change made")
                Return
            End If
            Try
                With USRstate.colBounds
                    .Remove(intDimension)
                    Dim colEntry As New Collection
                    With colEntry
                        .Add(intNewValue) : .Add(intUpperBound)
                    End With
                    .Add(colEntry, , intDimension)
                End With
            Catch
                errorHandler_("Error occured in modifying lower bound: " & _
                              Err.Number & " " & Err.Description, _
                              "LowerBound set", _
                              "Object is not usable")
                Me.mkUnusable()
                Return
            End Try
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Return a random array    
    '
    '
    Public Overloads Shared Function mkRandomArray() As String
        Return (mkRandomArray(SCALARTYPES & " VARIANT"))
    End Function
    Public Overloads Shared Function mkRandomArray(ByVal strScalarTypes As String) _
           As String
        Dim intExchange As Integer
        Dim intIndex1 As Integer
        Dim intLBound As Integer
        Dim intUBound As Integer
        Dim strDimensions As String
        For intIndex1 = 1 To Math.Max(1, CInt(Rnd() * 5))
            intLBound = CInt(Rnd() * 10) - 5
            intUBound = CInt(Rnd() * 10) - 5
            If intLBound > intUBound Then
                intExchange = intUBound
                intUBound = intLBound
                intLBound = intExchange
            End If
            _OBJutilities.append(strDimensions, _
                                 ",", _
                                 CStr(intLBound) & "," & CStr(intUBound))
        Next intIndex1
        Return ("Array," & _
               CStr(IIf(Rnd() < 0.75, mkRandomScalar, "Variant")) & "," & _
               strDimensions)
    End Function

    ' ----------------------------------------------------------------------
    ' Make a random domain
    '
    '
    ' --- Any domain
    Public Overloads Shared Function mkRandomDomain() As ENUvarType
        Return (mkRandomDomain(False))
    End Function
    ' --- Option restricts to a scalar domain
    Public Overloads Shared Function mkRandomDomain(ByVal booScalar As Boolean) As ENUvarType
        If booScalar Then
            Return (randomType_(False, SCALARTYPES))
        Else
            Return (randomType_(True, VARIABLETYPES))
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Make a random Quick Basic scalar as the fromString
    '
    '
    Public Overloads Shared Function mkRandomScalar() As String
        Return (mkRandomScalar(SCALARTYPES))
    End Function
    Public Overloads Shared Function mkRandomScalar(ByVal strTypes As String) _
           As String
        Return randomType_(False, strTypes).ToString
    End Function

    ' ----------------------------------------------------------------------
    ' Make a random Quick Basic scalar as the .Net value
    '
    '
    Public Overloads Shared Function mkRandomScalarValue() As Object
        Return (mkRandomScalarValue(SCALARTYPES))
    End Function
    Public Overloads Shared Function mkRandomScalarValue(ByVal strTypes As String) _
           As Object
        Dim strType As String = mkRandomScalar(strTypes)
        Select Case vtPrefixRemove(UCase(strType))
            Case "BOOLEAN" : Return (CBool(Rnd() > 0.5))
            Case "BYTE" : Return (CByte(Rnd() * 255))
            Case "INTEGER" : Return (CShort(Math.Max(Math.Min(CInt(Rnd() * 32768) _
                                                            * _
                                                            CInt(IIf(Rnd() > 0.5, 1, -1)), _
                                                            32767), _
                                                   -32768)))
            Case "LONG" : Return (CLng(Math.Max(Math.Min(CInt(Rnd() * (2 ^ 31 - 1)) _
                                                       * _
                                                       CInt(IIf(Rnd() > 0.5, 1, -1)), _
                                                       (2 ^ 31 - 1)), _
                                              -(2 ^ 31 - 1))))
            Case "SINGLE" : Return (CSng(Rnd() * 1000000 * CInt(IIf(Rnd() > 0.5, 1, -1))))
            Case "DOUBLE" : Return (CDbl(Rnd() * 1000000 * CInt(IIf(Rnd() > 0.5, 1, -1))))
            Case "STRING"
                Dim intIndex1 As Integer
                Dim strRandom As String = ""
                For intIndex1 = 1 To CInt(Rnd() * 10)
                    strRandom &= Mid("abcdefg", CInt(Rnd() * 5) + 1, 1)
                Next intIndex1
                Return (strRandom)
            Case Else
                errorHandler_("Internal programming error: invalid type " & _
                              _OBJutilities.enquote(strType) & " " & _
                              "was returned from mkRandomScalar", _
                              ClassName, _
                              "mkRandomScalarValue", _
                              "Returning a null string")
                Return ("")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Make a random type
    '
    '
    ' --- Public interface
    Public Shared Function mkRandomType() As String
        Dim dblRnd As Double = Rnd()
        If dblRnd <= 0.2 Then
            Return (mkRandomScalar())
        ElseIf dblRnd <= 0.4 Then
            Return (mkRandomArray())
        ElseIf dblRnd <= 0.6 Then
            Return (mkRandomVariant())
        ElseIf dblRnd <= 0.8 Then
            Return (mkRandomUDT())
        ElseIf dblRnd <= 0.9 Then
            Return ("Null")
        Else
            Return ("Unknown")
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Create a random UDT as the fromString
    '
    '
    Public Overloads Shared Function mkRandomUDT() As String
        Return (mkRandomUDT(0))
    End Function
    Private Overloads Shared Function mkRandomUDT(ByVal intLevel As Integer) _
            As String
        Dim intIndex1 As Integer
        Dim intRandom As Integer
        Dim strNext As String
        Dim strOutstring As String = "UDT"
        For intIndex1 = 1 To CInt(Rnd() * 4) + 1
            Select Case CInt(Rnd() * 3)
                Case 0 : strNext = mkRandomScalar()
                Case 1 : strNext = mkRandomArray()
                Case 2 : strNext = mkRandomVariant()
                Case 3
                    If intLevel > 5 Then
                        strNext = mkRandomScalar()
                    Else
                        strNext = mkRandomUDT(intLevel + 1)
                    End If
            End Select
            strOutstring &= ",(Member" & intIndex1 & "," & strNext & ")"
        Next intIndex1
        Return (strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Return a random variant containing a random array or scalar
    '
    '
    Public Shared Function mkRandomVariant() As String
        Dim strOutstring As String = "Variant,"
        If Rnd() <= 0.75 Then
            Return (strOutstring & mkRandomScalar())
        Else
            Return (strOutstring & "(" & mkRandomArray() & ")")
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Return a random variant value
    '
    '
    Public Shared Function mkRandomVariantValue() As String
        Return _
        (_OBJutilities.object2String(mkRandomScalarValue(randomType_(False, _
                                                                     SCALARTYPES).ToString)))
    End Function

    ' ----------------------------------------------------------------------
    ' Create the type from an enumerator
    '
    '
    Public Shared Function mkType(ByVal enuType As ENUvarType) As qbVariableType
        If enuType = ENUvarType.vtArray Then
            errorHandler_("Cannot make a type object from an abstract array", _
                          ClassName, _
                          "mkType", _
                          "Returning Nothing")
            Return (Nothing)
        End If
        Dim booOK As Boolean
        Dim intErr As Integer
        Dim objNew As qbVariableType
        Dim strErr As String
        Try
            objNew = New qbVariableType
            booOK = objNew.fromString(enuType.ToString)
        Catch
            intErr = Err.Number : strErr = Err.Description
        End Try
        If Not booOK Then
            errorHandler_("Cannot make the type object" & _
                          CStr(IIf(intErr = 0, "", intErr & " " & strErr)), _
                          ClassName, _
                          "mkType", _
                          "Returning Nothing")
            If Not (objNew Is Nothing) Then objNew.dispose()
            Return (Nothing)
        End If
        Return (objNew)
    End Function

    ' ----------------------------------------------------------------------
    ' Make the object unusable
    '
    '
    Public Function mkUnusable() As Boolean
        USRstate.booUsable = False : Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Return and change the object's name
    '
    '
    Public Overridable Property Name() As String
        Get
            Return (USRstate.strName)
        End Get
        Set(ByVal strNewValue As String)
            USRstate.strName = strNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Object constructor
    '
    '
    ' --- Create the Unknown data type
    Public Sub New()
        New_("Unknown")
    End Sub
    ' --- Identify the specifics in a fromString
    Public Sub New(ByVal strFromString As String)
        New_(strFromString)
    End Sub
    ' --- Create a shell data type that doesn't contain any data type info
    Private Sub New(ByVal booNothing As Boolean)
        New_("")
    End Sub
    Private Sub New_(ByVal strFromString As String)
        With USRstate
            Interlocked.Increment(_INTsequence)
            .strName = "qbVariableType" & _
                       _OBJutilities.alignRight(CStr(_INTsequence), 4, "0") & " " & _
                       Now
            Try
                OBJcollectionUtilities = New collectionUtilities.collectionUtilities
            Catch
                errorHandler_("Unable to make a collectionUtilities object: " & _
                              Err.Number & " " & Err.Description, _
                              "New", _
                              "Object is not usable")
                Return
            End Try
            createBoundList_()
            .booUsable = True
            If strFromString <> "" AndAlso Not Me.fromString(strFromString) Then
                .booUsable = False : Return
            End If
            .booUsable = inspection_(True)
        End With
    End Sub

    ' -----------------------------------------------------------------------
    ' Convert System name to generic type name
    '
    '
    Public Shared Function name2NetType(ByVal strSystemName As String) As String
        Dim booBoolean As Boolean
        Dim bytByte As Byte
        Dim shrShort As Short
        Dim intInteger As Integer
        Dim lngLong As Long
        Dim sglSingle As Single
        Dim dblDouble As Double
        Dim strString As String = ""
        Dim objObject As Object = ""
        Dim strSystemNameWork As String = UCase(strSystemName)
        If Len(strSystemNameWork) < 7 OrElse Mid(strSystemNameWork, 1, 7) <> "SYSTEM." Then
            strSystemNameWork = "SYSTEM." & strSystemNameWork
        End If
        Select Case Trim(UCase(strSystemNameWork))
            Case UCase(booBoolean.GetType.ToString) : Return ("BOOLEAN")
            Case UCase(bytByte.GetType.ToString) : Return ("BYTE")
            Case UCase(shrShort.GetType.ToString) : Return ("SHORT")
            Case UCase(intInteger.GetType.ToString) : Return ("INTEGER")
            Case UCase(lngLong.GetType.ToString) : Return ("INTEGER")
            Case UCase(sglSingle.GetType.ToString) : Return ("SINGLE")
            Case UCase(dblDouble.GetType.ToString) : Return ("DOUBLE")
            Case UCase(strString.GetType.ToString) : Return ("STRING")
            Case UCase(objObject.GetType.ToString) : Return ("OBJECT")
            Case Else
                Return (strSystemName)
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Convert generic net type to System name
    '
    '
    Public Shared Function netType2Name(ByVal strNetType As String) As String
        Dim booBoolean As Boolean
        Dim bytByte As Byte
        Dim shrShort As Short
        Dim intInteger As Integer
        Dim sglSingle As Single
        Dim dblDouble As Double
        Dim strString As String = ""
        Dim objObject As Object = ""
        Select Case Trim(UCase(strNetType))
            Case "BOOLEAN" : Return (booBoolean.GetType.ToString)
            Case "BYTE" : Return (bytByte.GetType.ToString)
            Case "SHORT" : Return (shrShort.GetType.ToString)
            Case "INTEGER" : Return (intInteger.GetType.ToString)
            Case "SINGLE" : Return (sglSingle.GetType.ToString)
            Case "DOUBLE" : Return (dblDouble.GetType.ToString)
            Case "STRING" : Return (strString.GetType.ToString)
            Case "OBJECT" : Return (objObject.GetType.ToString)
            Case Else
                errorHandler_("Invalid generic .Net type name " & _OBJutilities.enquote(strNetType), _
                              ClassName, _
                              "", _
                              "Returning a null string")
                Return ("")
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the type of the .Net scalar to the Quick Basic domain
    '
    '
    Public Shared Function netType2QBdomain(ByVal strType As String) As ENUvarType
        Select Case UCase(type2NetName_(strType))
            Case "BOOLEAN" : Return (ENUvarType.vtBoolean)
            Case "BYTE" : Return (ENUvarType.vtByte)
            Case "SHORT" : Return (ENUvarType.vtInteger)
            Case "INTEGER" : Return (ENUvarType.vtLong)
            Case "SINGLE" : Return (ENUvarType.vtSingle)
            Case "DOUBLE" : Return (ENUvarType.vtDouble)
            Case "STRING" : Return (ENUvarType.vtString)
            Case Else : Return (ENUvarType.vtUnknown)
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Convert .Net value to Quick Basic domain
    '
    '
    Public Shared Function netValue2QBdomain(ByVal objValue As Object) As ENUvarType
        If (objValue Is Nothing) Then
            Return ENUvarType.vtNull
        ElseIf (TypeOf objValue Is Boolean) Then
            Return ENUvarType.vtBoolean
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "BYTE") Is Nothing) Then
            Return ENUvarType.vtByte
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "SHORT") Is Nothing) Then
            Return ENUvarType.vtInteger
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "INTEGER") Is Nothing) Then
            Return ENUvarType.vtLong
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "SINGLE") Is Nothing) Then
            Return ENUvarType.vtSingle
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "DOUBLE") Is Nothing) Then
            Return ENUvarType.vtDouble
        ElseIf Not (_OBJutilities.canonicalTypeCast(objValue, "STRING") Is Nothing) Then
            Return ENUvarType.vtString
        Else
            Return ENUvarType.vtUnknown
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Convert .Net value to Quick Basic value
    '
    '
    ' --- Converts to narrowest type
    Public Shared Function netValue2QBvalue(ByVal objValue As Object) As Object
        Return netValue2QBvalue(objValue, netValue2QBdomain(objValue))
    End Function
    ' --- Converts to type specified as enum
    Public Shared Function netValue2QBvalue(ByVal objValue As Object, _
                                            ByVal enuType As ENUvarType) As Object
        Return netValue2QBvalue(objValue, enuType.ToString)
    End Function
    ' --- Converts to type specified as string
    Public Shared Function netValue2QBvalue(ByVal objValue As Object, _
                                            ByVal strType As String) As Object
        Try
            Select Case vtPrefixAdd(UCase(Trim(strType)))
                Case "VTBOOLEAN" : Return (CBool(objValue))
                Case "VTBYTE" : Return (CByte(objValue))
                Case "VTINTEGER" : Return (CShort(objValue))
                Case "VTLONG" : Return (CInt(objValue))
                Case "VTSINGLE" : Return (CSng(objValue))
                Case "VTDOUBLE" : Return (CDbl(objValue))
                Case "VTSTRING"
#If QUICKBASICENGINE_EXTENSION Then
                        Return(CStr(objValue))
#Else
                        Return (Mid(CStr(objValue), 1, CInt(2 ^ 16)))
#End If
                Case Else
                    errorHandler_("Unsupported type " & strType, _
                                  ClassName, _
                                  "netValue2QBvalue", _
                                  "Returning unchanged value with unchanged type")
            End Select
        Catch
            errorHandler_("Cannot convert .Net value " & _
                          _OBJutilities.object2String(objValue), _
                          "to Quick Basic value with type " & strType, _
                          "", _
                          Err.Number & " " & Err.Description)
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Test .Net value to make sure it is in the Quick Basic domain
    '
    '
    Public Overloads Function netValueInQBdomain(ByVal objValue As Object) As Boolean
        If Not checkUsable_("valueInDomain") Then Return (Nothing)
        Return netValueInQBdomain(USRstate.enuVariableType, objValue)
    End Function
    Public Overloads Shared Function netValueInQBdomain(ByVal enuType As ENUvarType, _
                                                        ByVal objValue As Object) _
           As Boolean
        Dim strNetType As String
        strNetType = UCase(objValue.GetType.ToString)
        Select Case enuType
            Case ENUvarType.vtArray
                Return (False)
            Case ENUvarType.vtBoolean
                Return (_OBJutilities.findWord(UCase(numericNetTypes_), strNetType) <> 0)
            Case ENUvarType.vtByte
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "BYTE") Is Nothing))
            Case ENUvarType.vtInteger
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "SHORT") Is Nothing))
            Case ENUvarType.vtLong
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "INTEGER") Is Nothing))
            Case ENUvarType.vtSingle
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "SINGLE") Is Nothing))
            Case ENUvarType.vtDouble
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "SINGLE") Is Nothing))
            Case ENUvarType.vtString
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "STRING") Is Nothing))
            Case ENUvarType.vtVariant
                Return (Not (_OBJutilities.canonicalTypeCast(objValue, "STRING") Is Nothing) _
                       OrElse _
                       (objValue Is Nothing))
            Case ENUvarType.vtNull
                Return (objValue Is Nothing)
            Case ENUvarType.vtUnknown
                Return (True)
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Tell whether .Net value can represent a QB scalar
    '
    '
    Public Shared Function netValueIsScalar(ByVal objValue As Object) As Boolean
        Select Case netValue2QBdomain(objValue)
            Case ENUvarType.vtArray : Return (False)
            Case ENUvarType.vtNull : Return (False)
            Case ENUvarType.vtUnknown : Return (False)
        End Select
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Convert .Net object to the corresponding type as a fromString
    '
    '
    Public Shared Function object2Type(ByVal objValue As Object) As String
        Select Case UCase(objValue.GetType.ToString)
            Case "SYSTEM.BOOLEAN" : Return "Boolean"
            Case "SYSTEM.BYTE" : Return "Byte"
            Case "SYSTEM.INT16" : Return "Integer"
            Case "SYSTEM.INT32" : Return "Long"
            Case "SYSTEM.SINGLE" : Return "Single"
            Case "SYSTEM.DOUBLE" : Return "Double"
            Case "SYSTEM.STRING" : Return "String"
            Case "SYSTEM.INT64"
                Try
                    Dim intValue As Integer = CInt(objValue)
                Catch
                    Return "Unknown"
                End Try
                Return "Long"
            Case Else : Return "Unknown"
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the object to XML
    '
    '
    Public Overloads Function object2XML(Optional ByVal booIncludeCache As Boolean = True) _
           As String
        Dim strCacheInfo As String
        If booIncludeCache Then
            strCacheInfo = vbNewLine & vbNewLine & vbNewLine & _
                           "CACHE INFO" & vbNewLine & vbNewLine & _
                           object2XML_getCacheInfo_()
        End If
        Return (object2XML(Me.About & _
                           vbNewLine & vbNewLine & vbNewLine & _
                           "This instance represents the following " & _
                           "variable type: " & _
                           vbNewLine & vbNewLine & _
                           Me.toDescription & _
                           strCacheInfo))
    End Function
    Friend Overloads Function object2XML(ByVal strComment As String, _
                                         Optional ByVal booIncludeCache As Boolean = False) _
           As String
        With USRstate
            Return (_OBJutilities.objectInfo2XML(Me.ClassName, _
                                                strComment, _
                                                True, True, _
                                                "booUsable", _
                                                "Indicates the usability of the object", _
                                                CStr(USRstate.booUsable), _
                                                "strName", _
                                                "Identifies the object instance", _
                                                .strName, _
                                                "enuVariableType", _
                                                "Identifies the variable's type", _
                                                .enuVariableType.ToString, _
                                                "objVarType", _
                                                "Identifies the type of a contained variable", _
                                                object2XML__varType2String_(.objVarType), _
                                                "colBounds", _
                                                "Identifies the bounds of an array type", _
                                                OBJcollectionUtilities.collection2String(.colBounds, booReadable:=True), _
                                                "colTypeOrdering", _
                                                "Identifies type ordering", _
                                                OBJcollectionUtilities.collection2String(.colTypeOrdering, booReadable:=True), _
                                                "booContained", _
                                                "Identifies type containment", _
                                                object2XML_contained2String_, _
                                                "objTag", _
                                                "User's tag", _
                                                _OBJutilities.object2String(.objTag)))
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Return array information on behalf of object2XML
    '
    '
    Private Function object2XML_arrayInfo_() As String
        With Me
            If .Dimensions = 0 Then Return (_OBJutilities.mkXMLComment("Scalar has no array info"))
            Dim strReturn As String = _OBJutilities.mkXMLComment("Array information") & _
                                      vbNewLine & _
                                      _OBJutilities.mkXMLElement("ElementType", _
                                                                .VarType.ToString) & _
                                      vbNewLine & _
                                      _OBJutilities.mkXMLElement("Dimensions", _
                                                                CStr(.Dimensions))
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Dimensions
                strReturn &= vbNewLine & _
                             _OBJutilities.mkXMLElement("LBound" & intIndex1, _
                                                       CStr(.LowerBound(intIndex1))) & _
                             vbNewLine & _
                             _OBJutilities.mkXMLElement("UBound" & intIndex1, _
                                                       CStr(.UpperBound(intIndex1)))
            Next intIndex1
            If CType(.ContainedVariableType, qbVariableType).VariableTypeEnum _
               = _
               ENUvarType.vtVariant Then
                strReturn &= vbNewLine & vbNewLine & _
                             CType(.ContainedVariableType, qbVariableType).object2XML _
                             ("Array's variant type")
            End If
            Return (strReturn)
        End With
    End Function

    ' ------------------------------------------------------------------------
    ' Return containment array on behalf of object2XML
    '
    '
    Private Function object2XML_contained2String_() As String
        With USRstate
            Dim intUBound1 As Integer
            Try
                intUBound1 = UBound(.booContained)
            Catch
                Return ("Unallocated")
            End Try
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim strOutstring As String
            For intIndex1 = 1 To UBound(.booContained, 1)
                For intIndex2 = 1 To UBound(.booContained, 2)
                    strOutstring &= CStr(IIf(.booContained(intIndex1, intIndex2), "Y", "N"))
                Next intIndex2
                If intIndex1 < UBound(.booContained, 1) Then
                    strOutstring &= ", "
                End If
            Next intIndex1
            Return (_OBJutilities.soft2HardParagraph(strOutstring, 32))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Get cache info on behalf of object2XML
    '
    '
    Private Function object2XML_getCacheInfo_() As String
        Dim enuExistence As _ENUcacheExistence = cacheExistence_()
        Dim strOutstring As String = "A cache of recently parsed variable types is maintained " & _
                                     "to save time: here is the state of the cache." & _
                                     vbNewLine & vbNewLine & _
                                     "Cache status: " & enuExistence.ToString
        If enuExistence <> _ENUcacheExistence.available Then Return strOutstring
        SyncLock _OBJcache
            With _OBJcache
                Dim strContains As String
                Dim intIndex1 As Integer
                Dim strNext As String
                For intIndex1 = 1 To .colCache.Count
                    strNext = CType(.colCache.Item(intIndex1), qbVariableType).ToString
                    If Len(strContains) + Len(strNext) > 50 Then
                        strContains &= "..." : Exit For
                    End If
                    strContains &= CStr(IIf(strContains = "", "", ", ")) & _
                                   _OBJutilities.enquote(_OBJutilities.ellipsis(strNext, 16))
                Next intIndex1
                Return strOutstring & vbNewLine & _
                       "Cache maximum size: " & .intCacheLimit & vbNewLine & _
                       "Cache current size: " & .colCache.Count & vbNewLine & _
                       "Cache contains: " & strContains
            End With
        End SyncLock
    End Function

    ' ------------------------------------------------------------------------
    ' Return contained varType as a string on behalf of object2XML
    '
    '
    Private Function object2XML__varType2String_(ByVal objVarType As Object) _
            As String
        With USRstate
            If (objVarType Is Nothing) Then
                Return (Nothing)
            ElseIf (TypeOf objVarType Is qbVariableType) Then
                Return (CType(objVarType, qbVariableType).ToString)
            ElseIf (TypeOf objVarType Is Collection) Then
                Return (OBJcollectionUtilities.collection2String(CType(objVarType, Collection), _
                                                                 booReadable:=True, _
                                                                 strSeparator1:=vbNewLine, _
                                                                 strSeparator:=vbNewLine))
            Else
                errorHandler_("Unexpected type of objVarType: objVarType=" & _
                              _OBJutilities.object2String(objVarType, True), _
                              "object2XML__varType2String_", _
                              "Marking object unusable: returning Invalid")
                Me.mkUnusable() : Return ("Invalid")
            End If
        End With
    End Function

    ' ------------------------------------------------------------------------
    ' Return variant information on behalf of object2XML
    '
    '
    Private Function object2XML_variantInfo_() As String
        With Me
            If .VariableType <> ENUvarType.vtVariant Then
                Return ("")
            End If
            Dim strXML As String
            If .Abstract Then
                strXML = _OBJutilities.mkXMLComment(vbNewLine & _
                                                    _OBJutilities.string2Box _
                                                    ("Variant is abstract and contains no value") & _
                                                    vbNewLine)
            Else
                strXML = CType(.ContainedVariableType, qbVariableType).object2XML _
                          ("Variant's contained value")
            End If
            Return (vbNewLine & "    " & _
                   Replace(strXML, vbNewLine, vbNewLine & "     ") & _
                   vbNewLine & vbNewLine)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the quickBasic domain to a .Net type
    '
    '
    Public Shared Function qbDomain2NetType(ByVal enuDomain As ENUvarType) As String
        Select Case enuDomain
            Case ENUvarType.vtBoolean : Return (netType2Name("Boolean"))
            Case ENUvarType.vtByte : Return (netType2Name("Byte"))
            Case ENUvarType.vtInteger : Return (netType2Name("Short"))
            Case ENUvarType.vtLong : Return (netType2Name("Integer"))
            Case ENUvarType.vtSingle : Return (netType2Name("Single"))
            Case ENUvarType.vtDouble : Return (netType2Name("Double"))
            Case ENUvarType.vtString : Return (netType2Name("String"))
            Case ENUvarType.vtVariant : Return (netType2Name("Object"))
            Case Else
                errorHandler_("ENUvarType " & enuDomain.ToString & " " & _
                              "doesn't convert to any .Net type", _
                              ClassName, _
                              "qbDomain2NetType", _
                              "Returning a null string")
                Return ("")
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Change the array's dimensions
    '
    '
    Public Function redimension(ByVal intDimension As Integer, _
                                ByVal intLowerBound As Integer, _
                                ByVal intUpperBound As Integer) As Boolean
        If Not checkUsable_("redimension") Then
            Return False
        End If
        If intLowerBound > intUpperBound Then
            errorHandler_("Lower bound " & intLowerBound & " " & _
                          "is greater than upper bound " & intLowerBound, _
                          "redimension", _
                          "No change made")
            Return False
        End If
        With Me
            If intDimension < 1 OrElse intDimension > .Dimensions Then
                errorHandler_("Rank " & intDimension & " " & _
                              "is invalid", _
                              "redimension", _
                              "No change made")
                Return False
            End If
            .LowerBound(intDimension) = .UpperBound(intDimension)
            If intLowerBound < .LowerBound(intDimension) Then
                .LowerBound(intDimension) = intLowerBound
                .UpperBound(intDimension) = intUpperBound
            Else
                .UpperBound(intDimension) = intUpperBound
                .LowerBound(intDimension) = intLowerBound
            End If
            Return True
        End With
    End Function


    ' -----------------------------------------------------------------------
    ' Return default value for scalar
    '
    '
    Public Function scalarDefault() As Object
        If Not (checkUsable_("scalarDefault")) Then
            Return (Nothing)
        End If
        With Me
            If .isUnknown Then Return (Nothing)
            If .isNull Then Return (Nothing)
            If .isScalar Then Return (Me.defaultValue(Me.VariableType))
            If .isArray Then Return (Me.defaultValue(Me.VarType))
            If .isVariant Then
                If .Abstract Then Return (Nothing)
                Return (Me.defaultValue(Me.VarType))
            End If
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Return abstract storage space
    '
    '
    Public ReadOnly Property StorageSpace() As Integer
        Get
            If Not checkUsable_("Storage Space Get") Then Return (0)
            If Me.isScalar OrElse Me.isVariant OrElse Me.isNull Then
                Return (1)
            ElseIf Me.isArray Then
                With Me
                    Dim intIndex1 As Integer
                    Dim intStorageSpace As Integer = .BoundSize(.Dimensions)
                    For intIndex1 = .Dimensions - 1 To 1 Step -1
                        intStorageSpace *= .BoundSize(intIndex1)
                    Next intIndex1
                    Return (intStorageSpace)
                End With
            ElseIf Me.isUDT Then
                With Me
                    Dim intIndex1 As Integer
                    Dim intStorageSpace As Integer
                    For intIndex1 = 1 To .UDTmemberCount
                        intStorageSpace += Me.innerType(intIndex1).StorageSpace
                    Next intIndex1
                    Return (intStorageSpace)
                End With
            ElseIf Me.isUnknown Then
                Return (0)
            Else
                errorHandler_("Internal programming error: unrecognizable type", _
                              "StorageSpace get", _
                              "Marking object unusable: returning 0")
                Me.mkUnusable()
                Return (0)
            End If
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Convert the string to a variant type enumerator
    '
    '
    Public Shared Function string2enuVarType(ByVal strInstring As String) _
           As ENUvarType
        Dim strWorkstring As String = UCase(vtPrefixAdd(Trim(strInstring)))
        Select Case strWorkstring
            Case "VTBOOLEAN"
                Return (ENUvarType.vtBoolean)
            Case "VTBYTE"
                Return (ENUvarType.vtByte)
            Case "VTINTEGER"
                Return (ENUvarType.vtInteger)
            Case "VTLONG"
                Return (ENUvarType.vtLong)
            Case "VTSINGLE"
                Return (ENUvarType.vtSingle)
            Case "VTDOUBLE"
                Return (ENUvarType.vtDouble)
            Case "VTSTRING"
                Return (ENUvarType.vtString)
            Case "VTVARIANT"
                Return (ENUvarType.vtVariant)
            Case "VTARRAY"
                Return (ENUvarType.vtArray)
            Case "VTNULL"
                Return (ENUvarType.vtNull)
            Case "VTUDT"
                Return (ENUvarType.vtUDT)
            Case "VTUNKNOWN"
                Return (ENUvarType.vtUnknown)
            Case Else
                errorHandler_("Invalid variant type string " & _
                              _OBJutilities.enquote(strInstring), _
                              "qbVariableType (shared)", _
                              "string2enuVarType", _
                              "Returning Unknown")
                Return (ENUvarType.vtUnknown)
        End Select
    End Function

    ' -----------------------------------------------------------------------
    ' Return and assign the user's Tag
    '
    '
    Public Property Tag() As Object
        Get
            If Not checkUsable_("Tag Get") Then Return (Nothing)
            Return (USRstate.objTag)
        End Get
        Set(ByVal objNewValue As Object)
            If Not checkUsable_("Tag Set") Then Return
            USRstate.objTag = objNewValue
        End Set
    End Property

#If Not QBVARIABLETEST_NOTEST Then
    ' -----------------------------------------------------------------------
    ' Test the object (using a contained object)
    '
    '
    Public Function test(ByRef strReport As String) As Boolean
        If Not checkUsable_("test") Then
            strReport = "Object isn't usable: cannot test" : Return (False)
        End If
        Dim datStart As Date = Now
        test_raiseEvent_("Starting test", 1)
        Dim intErrorCount As Integer
            Try
            ' --- Create the test object
            test_raiseEvent_("Creating object", 1)
            strReport = "Self-test of " & _
                        Me.Name & " " & _
                        "at " & Now & _
                        vbNewLine & vbNewLine & _
                        "Creating internal test object"
            Dim objTest As qbVariableType
            Try
                objTest = New qbVariableType
            Catch
                strReport &= vbNewLine & _
                            "Can't create internal test object: " & _
                            Err.Number & " " & Err.Description
                Me.mkUnusable()
                test_raiseEvent_("Test failed", -1)
                Return (False)
            End Try
            objTest.Name = objTest.Name & " (test object)"
            test_raiseEvent_("Test object " & _
                            _OBJutilities.enquote(objTest.Name) & " " & _
                            "has been created", _
                            -1)
            ' --- Make many fromString tests to test the parser
            test_raiseEvent_("Testing the fromString parser and the object", 1)
            Dim intIndex1 As Integer
            Dim strSplit() As String
            Try
                strSplit = Split(TEST_FROMSTRING_LIST_BASE & _
                                 vbNewLine & _
                                 test_randomExtensions_(), _
                                 vbNewLine)
            Catch
                errorHandler_("Cannot split fromString test list: " & _
                            Err.Number & " " & Err.Description, _
                            "test", _
                            "Marking object not usable and aborting test")
                Me.mkUnusable()
                test_raiseEvent_("Test failed", -1)
                Return (False)
            End Try
            Dim strInspectionReport As String
            Dim strFromString As String
            Dim intErr As Integer
            Dim strDesc As String
            For intIndex1 = 0 To UBound(strSplit)
                test_raiseEvent_("Testing the fromString parser and the object", _
                                 "test case", _
                                 intIndex1 + 1, _
                                 UBound(strSplit) + 1)
                strReport &= vbNewLine & _
                            "Testing fromString(" & strSplit(intIndex1) & ")"
                intErr = 0
                Try
                    objTest.fromString(strSplit(intIndex1))
                    If Not objTest.inspect(strInspectionReport) Then
                        intErrorCount += 1
                        strReport &= vbNewLine & _
                                    strFromString & " has failed: " & _
                                    "inspection failed, report follows" & _
                                    vbNewLine & _
                                    _OBJutilities.string2Box(strInspectionReport, _
                                                            "Inspection report")
                    End If
                Catch
                    intErr = Err.Number
                    strDesc = Err.Description
                End Try
                strFromString = "fromString(" & strSplit(intIndex1) & ")"
                If intErr = 0 Then
                    strReport &= vbNewLine & _
                                strFromString & " " & "has succeeded"
                Else
                    intErrorCount += 1
                    strReport &= vbNewLine & _
                                strFromString & " " & "has failed: " & _
                                intErr & " " & strDesc
                End If
                strReport &= vbNewLine & vbNewLine & _
                            _OBJutilities.string2Box(objTest.object2XML, _
                                                     "XML dump of test object at " & Now)
            Next intIndex1
            test_raiseEvent_("Testing of the fromString parser and object creation is complete: " & _
                            "error count: " & intErrorCount, _
                            -1)
            ' --- Test default values
            test_raiseEvent_("Testing the default value method", 1)
            Try
                strSplit = Split(TEST_DEFAULTVALUE, vbNewLine)
            Catch
                errorHandler_("Cannot split defaultValue test list: " & _
                            Err.Number & " " & Err.Description, _
                            "test", _
                            "Making the object unusable and returning False")
                Me.mkUnusable()
                test_raiseEvent_("Test failed", -1)
                Return (False)
            End Try
            Dim strDefaultActual As String
            Dim strDefaultExpected As String
            For intIndex1 = 0 To UBound(strSplit)
                test_raiseEvent_("Testing the default value method", _
                                "scalar type", _
                                intIndex1 + 1, _
                                UBound(strSplit) + 1)
                strFromString = _OBJutilities.item(strSplit(intIndex1), 1, ",", False)
                If objTest.fromString(strFromString) Then
                    strDefaultActual = _OBJutilities.object2String(objTest.defaultValue, True)
                    strDefaultExpected = Trim(_OBJutilities.item(strSplit(intIndex1), 2, ",", False))
                    If strDefaultExpected <> "Nothing" Then
                        strDefaultExpected = netType2Name(_OBJutilities.item(strDefaultExpected, _
                                                                            1, _
                                                                            "(", _
                                                                            False)) & _
                                            "(" & _
                                            _OBJutilities.item(strDefaultExpected, 2, "(", False)
                    End If
                    If Trim(UCase(strDefaultActual)) _
                    = _
                    Trim(UCase(strDefaultExpected)) Then
                        strReport &= vbNewLine & vbNewLine & _
                                    "defaultValue correctly produces " & _
                                    _OBJutilities.enquote(strDefaultActual) & " " & _
                                    "for type " & strFromString
                    Else
                        strReport &= vbNewLine & vbNewLine & _
                                    "defaultValue produces unexpected value " & _
                                    _OBJutilities.enquote(strDefaultActual) & " " & _
                                    "for type " & strFromString & ": " & _
                                    "expected value is " & _
                                    _OBJutilities.enquote(strDefaultExpected)
                        intErrorCount += 1
                    End If
                Else
                    strReport &= vbNewLine & vbNewLine & _
                                "Cannot convert string " & _
                                _OBJutilities.enquote(strFromString) & " " & _
                                "from type"
                    intErrorCount += 1
                End If
            Next intIndex1
            test_raiseEvent_("Testing of defaults is complete: " & _
                            "number of errors: " & intErrorCount, _
                            -1)
            ' --- Test type containment
            test_raiseEvent_("Testing containedType", 1)
            strReport &= vbNewLine & vbNewLine & _
                         "Testing containedType" & _
                         vbNewLine & vbNewLine
            Try
                strSplit = Split(TEST_CONTAINMENT, vbNewLine)
            Catch
                errorHandler_("Cannot split containment list" & _
                            Err.Number & " " & Err.Description, _
                            "test", _
                            "Marking the object unusable and returning False")
                Me.mkUnusable()
                test_raiseEvent_("Test failed", -1)
                Return (False)
            End Try
            Dim booArray As Boolean
            Dim booContained As Boolean
            Dim booOK As Boolean
            Dim objCTtest(1) As qbVariableType
            Dim strSplit2() As String
            Try
                objCTtest(0) = New qbVariableType
                objCTtest(1) = New qbVariableType
            Catch
                errorHandler_("Cannot create test objects: " & _
                                Err.Number & " " & Err.Description, _
                                "", _
                                "Marking the object unusable and returning False")
                Me.mkUnusable()
                test_raiseEvent_("Test failed", -1)
                Return (False)
            End Try
            For intIndex1 = 0 To UBound(strSplit)
                test_raiseEvent_("Testing containedType", _
                                 "test case", _
                                 intIndex1 + 1, UBound(strSplit) + 1)
                Try
                    strSplit2 = Split(strSplit(intIndex1), ":")
                Catch
                    errorHandler_("Cannot split containment list entry" & _
                                  Err.Number & " " & Err.Description, _
                                  "test", _
                                  "Marking the object unusable and returning False")
                    Me.mkUnusable()
                    test_raiseEvent_("Test failed", -1)
                    Return (False)
                End Try
                booArray = False
                If UCase(Trim(strSplit2(0))) = "ARRAY" Then
                    strSplit2(0) = Me.mkRandomArray : booArray = True
                End If
                If UCase(Trim(strSplit2(1))) = "ARRAY" Then
                    strSplit2(1) = Me.mkRandomArray : booArray = True
                End If
                Dim intRnd As Integer
                If booArray Then
                    objCTtest(0).fromString(strSplit2(0))
                    objCTtest(1).fromString(strSplit2(1))
                    booContained = objTest.containedType(objCTtest(0), objCTtest(1))
                Else
                    intRnd = CInt(Rnd() * 7)
                    Select Case intRnd
                        Case 0
                            booContained = objTest.containedType(string2enuVarType(strSplit2(0)), _
                                                                 string2enuVarType(strSplit2(1)))
                        Case 1
                            objCTtest(0).fromString(strSplit2(1))
                            booContained = objTest.containedType(string2enuVarType(strSplit2(0)), _
                                                                 objCTtest(0))
                        Case 2
                            objCTtest(0).fromString(strSplit2(0))
                            booContained = objTest.containedType(objCTtest(0), _
                                                                 string2enuVarType(strSplit2(1)))
                        Case 3
                            objCTtest(0).fromString(strSplit2(0))
                            objCTtest(1).fromString(strSplit2(1))
                            booContained = objTest.containedType(objCTtest(0), objCTtest(1))
                        Case 4
                            booContained = objTest.containedTypeWithState(string2enuVarType(strSplit2(0)), _
                                                                          string2enuVarType(strSplit2(1)))
                        Case 5
                            objCTtest(0).fromString(strSplit2(1))
                            booContained = objTest.containedTypeWithState(string2enuVarType(strSplit2(0)), _
                                                                          objCTtest(0))
                        Case 6
                            objCTtest(0).fromString(strSplit2(0))
                            booContained = objTest.containedTypeWithState(objCTtest(0), _
                                                                           string2enuVarType(strSplit2(1)))
                        Case 7
                            objCTtest(0).fromString(strSplit2(0))
                            objCTtest(1).fromString(strSplit2(1))
                            booContained = objTest.containedTypeWithState(objCTtest(0), objCTtest(1))
                        Case Else
                            errorHandler_("Unexpected Case", _
                                        "test", _
                                        "Marking object unusable and aborting the test")
                            Me.mkUnusable() : Return (False)
                    End Select
                    booOK = (booContained = CBool(Trim(strSplit2(2))))
                    strReport &= vbNewLine & _
                                    "Test object " & _
                                    CStr(IIf(booOK, "", "in")) & _
                                    "correctly indicates that " & _
                                    strSplit2(0) & " " & _
                                    "is " & CStr(IIf(booContained, "", "not ")) & _
                                    "contained in " & _
                                    strSplit2(1)
                    If Not booOK Then intErrorCount += 1
                End If
            Next intIndex1
            objCTtest(0).dispose() : objCTtest(1).dispose()
            test_raiseEvent_("Completed testing containedType methods", -1)
            ' --- Test Net/Quick Basic domain methods
            test_raiseEvent_("Testing the Net/Quick Basic domain methods", 1)
            strReport &= vbNewLine & vbNewLine & _
                        "Testing qbDomain2NetType and netType2QBdomain"
            With objTest
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtBoolean)) <> ENUvarType.vtBoolean Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Boolean"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtByte)) <> ENUvarType.vtByte Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Byte"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtInteger)) <> ENUvarType.vtInteger Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Integer"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtLong)) <> ENUvarType.vtLong Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Long"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtSingle)) <> ENUvarType.vtSingle Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Single"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtDouble)) <> ENUvarType.vtDouble Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Double"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtString)) <> ENUvarType.vtString Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for String"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtBoolean)) <> ENUvarType.vtBoolean Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Boolean"
                    intErrorCount += 1
                End If
                If .netType2QBdomain(.qbDomain2NetType(ENUvarType.vtBoolean)) <> ENUvarType.vtBoolean Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netType2QBdomain and/or qbDomain2NetType fails for Boolean"
                    intErrorCount += 1
                End If
            End With
            strReport &= vbNewLine & vbNewLine & _
                        "Testing netType2QBdomain and netTypeInQBdomain"
            objTest.fromString("Integer")
            With objTest
                If .netValue2QBdomain(255) <> ENUvarType.vtByte Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Byte"
                    intErrorCount += 1
                End If
                If .netValue2QBdomain(-32768) <> ENUvarType.vtInteger Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Integer"
                    intErrorCount += 1
                End If
                If .netValue2QBdomain(2 ^ 31 - 1) <> ENUvarType.vtLong Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Long"
                    intErrorCount += 1
                End If
                If .netValue2QBdomain(CStr("aa")) <> ENUvarType.vtString Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for String"
                    intErrorCount += 1
                End If
                If .netValue2QBdomain(Nothing) <> ENUvarType.vtNull Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Null"
                    intErrorCount += 1
                End If
                .fromString("Boolean")
                If Not .netValueInQBdomain(CBool(True)) Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Boolean"
                    intErrorCount += 1
                End If
                If Not .netValueInQBdomain(CByte(0)) Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Byte"
                    intErrorCount += 1
                End If
                If Not .netValueInQBdomain(ENUvarType.vtInteger, CByte(0)) Then
                    strReport &= vbNewLine & vbNewLine & _
                                "netValue2QBdomain fails for Byte/Integer"
                    intErrorCount += 1
                End If
            End With
            test_raiseEvent_("Completed testing the Net/Quick Basic domain methods", -1)
        Catch objException As Exception
            strReport &= vbNewLine & vbNewLine & _
                         "The following error occured: " & _
                         Err.Number & " " & Err.Description & " " & _
                         vbNewLine & vbNewLine & _
                         objException.Message
            intErrorCount += 1
        End Try
        test_raiseEvent_("Completed the test", -1)
        If intErrorCount = 0 Then
            strReport &= vbNewLine & vbNewLine & _
                         "Test succeeded: test took " & _
                         DateDiff("s", datStart, Now) & " second(s)"
            test_raiseEvent_("Test succeeded", -1)
            Return (True)
        Else
            strReport &= vbNewLine & vbNewLine & _
                         "Test has failed with " & intErrorCount & " error(s) : object is not usable"
            Me.mkUnusable()
            test_raiseEvent_("Test failed", -1)
            Return (False)
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Raise fun events on behalf of the test method
    '
    '
    ' --- Raise test events (no level change)  
    Private Overloads Sub test_raiseEvent_(ByVal strDesc As String)
        test_raiseEvent__(strDesc, False, 0, "", 0, 0)
    End Sub
    ' --- Raise test events at new level
    Private Overloads Sub test_raiseEvent_(ByVal strDesc As String, _
                                           ByVal intLevelChange As Integer)
        test_raiseEvent__(strDesc, False, intLevelChange, "", 0, 0)
    End Sub
    ' --- Raise test progress events
    Private Overloads Sub test_raiseEvent_(ByVal strDesc As String, _
                                           ByVal strEntity As String, _
                                           ByVal intEntityNumber As Integer, _
                                           ByVal intEntityCount As Integer)
        test_raiseEvent__(strDesc, True, 0, "test case", intEntityNumber, intEntityCount)
    End Sub
    ' --- Common logic
    Private Sub test_raiseEvent__(ByVal strDesc As String, _
                                  ByVal booIsProgress As Boolean, _
                                  ByVal intLevelChange As Integer, _
                                  ByVal strEntity As String, _
                                  ByVal intEntityNumber As Integer, _
                                  ByVal intEntityCount As Integer)
        Dim strMessage As String
        If booIsProgress Then
            If intEntityCount <> 0 Then
                strMessage = ": " & _
                             Math.Round(intEntityNumber / intEntityCount * 100, 2) & "%"
            End If
            strMessage = "at " & strEntity & " " & _
                         intEntityNumber & " of " & intEntityCount & _
                         strMessage
        End If
        strMessage = "[" & Me.Name & "] " & Now & " " & _
                     strDesc & strMessage
        If booIsProgress Then
            RaiseEvent testProgressEvent(strDesc, _
                                         strEntity, _
                                         intEntityNumber, _
                                         intEntityCount)
        Else
            RaiseEvent testEvent(strDesc, intLevelChange)
        End If
    End Sub

    ' -----------------------------------------------------------------------
    ' Return additional cases on behalf of the test method
    '
    '
    Private Function test_randomExtensions_() As String
        Dim intIndex1 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To 25
            _OBJutilities.append(strOutstring, vbNewLine, mkRandomType)
        Next intIndex1
        Return (strOutstring)
    End Function

#End If

    ' -----------------------------------------------------------------------
    ' Report test availability
    '
    '
    Public Shared ReadOnly Property TestAvailable() As Boolean
        Get
#If QBVARIABLETEST_NOTEST Then
                Return(False)
#Else
                Return (True)
#End If
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Return contained type name for array or variant: return null for scalar
    '
    '
    ' --- Variant or array syntax
    Public Overloads Function toContainedName() As String
        If Not checkUsable_("toContainedName") Then Return ("")
        If Not Me.isArray And Not Me.isVariant Then Return ("")
        Return (CType(USRstate.objVarType, qbVariableType).toName)
    End Function
    ' --- UDT syntax   
    Public Overloads Function toContainedName(ByVal objIndex As Object) As String
        If Not checkUsable_("toContainedName") Then Return ("")
        If Not Me.isUDT Then
            errorHandler_("This variable type is not a UDT: index is not allowed", _
                          "toContainedName", _
                          "Returning a null string")
            Return ("")
        End If
        Dim intIndex1 As Integer = udtIndex2Integer_(objIndex, Me.UDTmemberCount)
        If intIndex1 = 0 Then Return ("")
        Return (CType(CType(CType(USRstate.objVarType, _
                                 Collection).Item(intIndex1), _
                           Collection).Item(3), _
                     qbVariableType).toName)
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the type to its description
    '
    '
    Public Function toDescription() As String
        If Not checkUsable_("toDescription") Then
            Return ("Type description is not available because " & _
                   "object is unusable")
        End If
        With USRstate
            Select Case USRstate.enuVariableType
                Case ENUvarType.vtBoolean : Return (TODESCRIPTION_BOOLEAN)
                Case ENUvarType.vtByte : Return (TODESCRIPTION_BYTE)
                Case ENUvarType.vtInteger : Return (TODESCRIPTION_INTEGER)
                Case ENUvarType.vtLong : Return (TODESCRIPTION_LONG)
                Case ENUvarType.vtSingle : Return (TODESCRIPTION_SINGLE)
                Case ENUvarType.vtDouble : Return (TODESCRIPTION_DOUBLE)
                Case ENUvarType.vtString
#If QBVARIABLETYPE_EXTENSION Then
                        Return(TODESCRIPTION_STRINGNOLIMIT)
#Else
                        Return (TODESCRIPTION_STRING)
#End If
                Case ENUvarType.vtVariant
                    If Me.Abstract Then
                        Return (TODESCRIPTION_ABSTRACTVARIANT)
                    End If
                    Return ("Variant containing " & CType(.objVarType, qbVariableType).toDescription)
                Case ENUvarType.vtArray
                    Select Case Me.Dimensions
                        Case 1
                            Return ("1-dimensional array containing " & _
                                   Me.BoundSize(1) & " " & _
                                   "elements from " & _
                                   Me.LowerBound(1) & " to " & Me.UpperBound(1) & ": " & _
                                   "each element has the type " & _
                                   CType(.objVarType, qbVariableType).toDescription) & ": " & _
                                   "total size is " & _
                                   Me.StorageSpace
                        Case 2
                            Return ("2-dimensional array with " & _
                                   Me.BoundSize(1) & " " & _
                                   "rows " & _
                                   "(from " & _
                                   Me.LowerBound(1) & " to " & Me.UpperBound(1) & ") and " & _
                                   Me.BoundSize(2) & " columns: " & _
                                   "(from " & _
                                   Me.LowerBound(2) & " to " & Me.UpperBound(2) & "): " & _
                                   "each element has the type " & _
                                   CType(.objVarType, qbVariableType).toDescription) & ": " & _
                                   "total size is " & _
                                   Me.StorageSpace
                        Case Else
                            Return (Me.Dimensions & "-dimensional array with these bounds: " & _
                                   _OBJutilities.itemPhrase(Me.ToString, 3, -1, ",", False) & ": " & _
                                   "each element has the type " & _
                                   CType(.objVarType, qbVariableType).toDescription) & ": " & _
                                   "total size is " & _
                                   Me.StorageSpace
                    End Select
                Case ENUvarType.vtUnknown : Return (Me.ToString & _
                                                  ": represents an unknown type and/or value")
                Case ENUvarType.vtNull : Return (Me.ToString & _
                                               ": represents a null type and/or value")
                Case ENUvarType.vtUDT
                    With CType(.objVarType, Collection)
                        Dim colHandle As Collection
                        Dim intIndex1 As Integer
                        Dim strOutstring As String = "Type:"
                        For intIndex1 = 1 To Me.UDTmemberCount
                            colHandle = CType(.Item(intIndex1), Collection)
                            With colHandle
                                _OBJutilities.append(strOutstring, _
                                                     vbNewLine, _
                                                     CStr(.Item(2)) & ": " & _
                                                     CType(.Item(3), qbVariableType).toDescription)
                            End With
                        Next intIndex1
                        Return (strOutstring & _
                               vbNewLine & _
                               "End Type: total size is " & Me.StorageSpace)
                    End With
                Case Else
                    Dim strError As String = _
                        "Internal programming error: unexpected variable type"
                    errorHandler_("Internal programming error: unexpected variable type", _
                                  "toDescription", _
                                  "Marking object unusable and returning error info")
                    Me.mkUnusable()
                    Return (strError)
            End Select
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the type to its name
    '
    '
    Public Function toName() As String
        If Not checkUsable_("toName") Then Return ("")
        Return Me.VariableType.ToString
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the type to a string
    '
    '
    Public Overrides Function toString() As String
        Return (toStringVerify(booVerify:=False))
    End Function
    Public Function toStringVerify(Optional ByVal booVerify As Boolean = False) _
           As String
        If Not checkUsable_("toString") Then Return ("Unusable object instance")
        Try
            With USRstate
                Dim strToString As String = vtPrefixRemove(.enuVariableType.ToString)
                If isScalarType(.enuVariableType) _
                   OrElse _
                   .enuVariableType = ENUvarType.vtNull _
                   OrElse _
                   .enuVariableType = ENUvarType.vtUnknown Then Return (strToString)
                If Not .objVarType Is Nothing Then
                    strToString = strToString & ","
                    Dim strToStringInner As String
                    If (TypeOf .objVarType Is qbVariableType) Then
                        strToStringInner = CType(.objVarType, qbVariableType).ToString
                        Dim booComma As Boolean = (InStr(strToStringInner, ",") <> 0)
                        strToStringInner = CStr(IIf(booComma, "(", "")) & _
                                           strToStringInner & _
                                           CStr(IIf(booComma, ")", ""))
                    ElseIf (TypeOf .objVarType Is Collection) Then
                        With CType(.objVarType, Collection)
                            Dim colHandle As Collection
                            Dim intIndex1 As Integer
                            For intIndex1 = 1 To .Count
                                colHandle = CType(.Item(intIndex1), Collection)
                                With colHandle
                                    _OBJutilities.append(strToStringInner, _
                                                         ",", _
                                                         "(" & _
                                                         CStr(.Item(2)) & "," & _
                                                         CType(.Item(3), qbVariableType).ToString & _
                                                         ")")
                                End With
                            Next intIndex1
                        End With
                    Else
                        errorHandler_("Invalid type found", _
                                      "toStringVerify", _
                                      "Marking object unusable and returning Invalid")
                        Me.mkUnusable() : strToStringInner = "Invalid"
                    End If
                    strToString = strToString & strToStringInner
                End If
                If .enuVariableType = ENUvarType.vtVariant Then
                    Return (strToString)
                End If
                If .enuVariableType = ENUvarType.vtArray Then
                    strToString &= "," & toString_dimensionList_()
                End If
                If booVerify AndAlso Not checkToFromString_(strToString) Then
                    _OBJutilities.errorHandler("toString value " & _
                                                _OBJutilities.enquote(strToString) & " " & _
                                                "does not create identical object " & _
                                                "using the fromString method")
                End If
                Return (strToString)
            End With
        Catch
            Return ("Cannot create toString: " & Err.Number & " " & Err.Description)
        End Try
    End Function

    ' ------------------------------------------------------------------------
    ' Convert the dimension list to a string
    '
    '
    Private Function toString_dimensionList_() As String
        With USRstate
            Dim strDimensionList As String
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .colBounds.Count
                With CType(.colBounds.Item(intIndex1), Collection)
                    _OBJutilities.append(strDimensionList, _
                                        ",", _
                                        CStr(.Item(1)) & "," & CStr(.Item(2)))
                End With
            Next intIndex1
            Return (strDimensionList)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return the number of UDT members
    '
    '
    Public ReadOnly Property UDTmemberCount() As Integer
        Get
            If Not checkUsable_("UDTmemberCount Get") Then Return (0)
            If Not Me.isUDT Then
                errorHandler_("Object is not a UDT", _
                              "UDTmemberCount Get", _
                              "Returning 0")
            End If
            Return (CType(USRstate.objVarType, Collection).Count)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Converts member name to member index
    '
    '
    Public Function udtMemberDeref(ByVal strMemberName As String) As Integer
        If Not checkUsable_("strMemberName") Then Return (0)
        Return (udtIndex2Integer_(strMemberName, Me.UDTmemberCount))
    End Function

    ' -----------------------------------------------------------------------
    ' Return UDT member name
    '
    '
    Public Function udtMemberName(ByVal objIndex As Object) As String
        If Not checkUsable_("udtMemberName") Then Return ("")
        If Not Me.isUDT Then
            errorHandler_("Object is not a UDT", "udtMemberName", "Returning null string")
            Return ("")
        End If
        Dim intIndex1 As Integer = udtIndex2Integer_(objIndex, Me.UDTmemberCount)
        If intIndex1 = 0 Then Return ("")
        Return (CStr(CType(CType(USRstate.objVarType, Collection).Item(intIndex1), _
                          Collection).Item(2)))
    End Function

    ' ----------------------------------------------------------------------
    ' Return the upperBound at the indicated dimension
    '
    '
    Public Property UpperBound(ByVal intDimension As Integer) As Integer
        Get
            If Not checkUsable_("UpperBound get") Then Return (0)
            If Me.VariableType <> ENUvarType.vtArray Then
                errorHandler_("Upper bound requested for nonarray", _
                              "UpperBound get", _
                              "Returning 0")
                Return (0)
            End If
            If intDimension < 1 OrElse intDimension > Me.Dimensions Then
                errorHandler_("Dimension requested for upper bound is less than 1 " & _
                              "or greater than dimensions " & Me.Dimensions, _
                              "UpperBound get", _
                              "Returning 0")
                Return (0)
            End If
            Return (CInt(CType(USRstate.colBounds.Item(intDimension), Collection).Item(2)))
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("UpperBound set") Then Return
            If Me.VariableType <> ENUvarType.vtArray Then
                errorHandler_("Upper bound requested for nonarray", _
                              "UpperBound set", _
                              "No change made")
                Return
            End If
            If intDimension < 1 OrElse intDimension > Me.Dimensions Then
                errorHandler_("Dimension requested for upper bound is less than 1 " & _
                              "or greater than dimensions " & Me.Dimensions, _
                              "UpperBound set", _
                              "No change made")
                Return
            End If
            Dim intLowerBound As Integer = Me.LowerBound(intDimension)
            If intNewValue < intLowerBound Then
                errorHandler_("Proposed upper bound " & intNewValue & " " & _
                              "is less than actual lower bound" & _
                              intLowerBound, _
                              "UpperBound set", _
                              "No change made")
                Return
            End If
            Try
                With USRstate.colBounds
                    .Remove(intDimension)
                    Dim colEntry As New Collection
                    With colEntry
                        .Add(intLowerBound) : .Add(intNewValue)
                    End With
                    .Add(colEntry, , intDimension)
                End With
            Catch
                errorHandler_("Error occured in modifying lower bound: " & _
                              Err.Number & " " & Err.Description, _
                              "LowerBound set", _
                              "Object is not usable")
                Me.mkUnusable()
                Return
            End Try
        End Set
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
    ' Test the fromString for validity
    '
    '
    ' --- Check validity and destroy test object
    Public Overloads Shared Function validFromString(ByVal strFS As String) As Boolean
        Dim objTest As qbVariableType
        If Not validFromString(strFS, objTest) Then Return False
        objTest.dispose()
        Return True
    End Function
    ' --- Check validity and return test object
    Public Overloads Shared Function validFromString(ByVal strFS As String, _
                                                     ByRef objTest As qbVariableType) As Boolean
        Try
            objTest = New qbVariableType(strFS)
        Catch
            Return False
        End Try
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' Return the type
    '
    '
    Public ReadOnly Property VariableType() As ENUvarType
        Get
            If Not checkUsable_("VariableType Get") Then Return (ENUvarType.vtUnknown)
            Return (USRstate.enuVariableType)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Return a list of variable types
    '
    '
    Public Shared ReadOnly Property VariableTypeList() As String
        Get
            Return (VARIABLETYPES)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Return the type of the contained variable
    '
    '
    Public ReadOnly Property VarType() As ENUvarType
        Get
            If Not checkUsable_("VarType Get") Then Return (ENUvarType.vtUnknown)
            With USRstate
                If (.objVarType Is Nothing) OrElse Me.isUDT Then
                    Return (ENUvarType.vtUnknown)
                End If
                Return (CType(.objVarType, qbVariableType).VariableType)
            End With
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Conditionally, add prefix (vt) to variable type name  
    '
    '
    Public Shared Function vtPrefixAdd(ByVal strName As String) As String
        If LCase(Mid(strName, 1, 2)) <> "vt" Then Return ("vt" & strName)
        Return (strName)
    End Function

    ' ----------------------------------------------------------------------
    ' Conditionally, remove prefix (vt) to variable type name  
    '
    '
    Public Shared Function vtPrefixRemove(ByVal strName As String) As String
        If LCase(Mid(strName, 1, 2)) <> "vt" Then Return (strName)
        Return (Mid(strName, 3))
    End Function

    ' ***** Friend procedures ***********************************************

    ' -----------------------------------------------------------------------
    ' Change variable type
    '
    '
    Friend Function changeVariableType(ByVal enuType As ENUvarType) As Boolean
        If Not checkUsable_("changeVariableType") Then Return (False)
        USRstate.enuVariableType = enuType
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Assign and return the contained variable type
    '
    '
    Friend Property ContainedVariableType() As Object
        Get
            Return (USRstate.objVarType)
        End Get
        Set(ByVal objNewValue As Object)
            USRstate.objVarType = objNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Assign and return foreign state
    '
    '
    Friend Property State() As TYPstate
        Get
            Return (USRstate)
        End Get
        Set(ByVal typNewState As TYPstate)
            With USRstate
                .booContained = typNewState.booContained
                .booUsable = typNewState.booUsable
                .colBounds = OBJcollectionUtilities.collectionCopy(typNewState.colBounds)
                Try
                    .colTypeOrdering = OBJcollectionUtilities.collectionCopy(typNewState.colTypeOrdering)
                Catch : End Try
                .enuVariableType = typNewState.enuVariableType
                .objTag = typNewState.objTag
                If (typNewState.objVarType Is Nothing) Then
                    .objVarType = Nothing
                ElseIf (TypeOf typNewState.objVarType Is qbVariableType) Then
                    Try
                        .objVarType = New qbVariableType
                    Catch
                        errorHandler_("Programming error: cannot create new var type: " & _
                                      Err.Number & " " & Err.Description, _
                                      "State Set", _
                                      "Marking object not usable")
                        mkUnusable()
                        Return
                    End Try
                    CType(.objVarType, qbVariableType).State = _
                    CType(typNewState.objVarType, qbVariableType).State
                ElseIf (TypeOf typNewState.objVarType Is Collection) Then
                    .objVarType = _
                    OBJcollectionUtilities.collectionCopy(CType(typNewState.objVarType, Collection))
                Else
                    errorHandler_("Programming error: objVarType in new state " & _
                                  "has unsupported type", _
                                  "State Set", _
                                  "Marking object not usable")
                    mkUnusable()
                    Return
                End If
                .strName = typNewState.strName
            End With
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Assign and return the variable type
    '
    '
    Friend Property VariableTypeEnum() As ENUvarType
        Get
            Return (USRstate.enuVariableType)
        End Get
        Set(ByVal enuNewValue As ENUvarType)
            USRstate.enuVariableType = enuNewValue
        End Set
    End Property

    ' ***** Private procedures **********************************************

    ' -----------------------------------------------------------------------
    ' Check the cache for the fromString expression and a pre-existing
    ' qbVariableType, we can clone and reuse
    '
    '
    Private Function cacheCheck_(ByVal strFromString As String) As Boolean
        If cacheExistence_() <> _ENUcacheExistence.available Then Return False
        SyncLock _OBJcache
            Dim objCached As qbVariableType
            Try
                objCached = CType(_OBJcache.colCache.Item(strFromString), qbVariableType)
            Catch
                Return False
            End Try
            Dim strName As String = Me.Name
            Me.State = objCached.State
            Me.Name = strName
        End SyncLock
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' Make shared cache
    '
    '
    ' Note that after the exchange and before the code under Try is complete,
    ' the cache will be Nothing and its existence will be "beingCreated".
    '
    Private Function cacheCreate_() As Boolean
        Interlocked.Exchange(_INTcacheExistence, 1)
        Try
            _OBJcache = New _cache
            SyncLock _OBJcache
                With _OBJcache
                    .intCacheLimit = DEFAULT_CACHE_LIMIT
                    .colCache = New Collection
                End With
            End SyncLock
            Return True
        Catch
            Return False
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Return True if the cache exists, false otherwise
    '
    '
    Private Shared Function cacheExistence_() As _ENUcacheExistence
        Dim booExistence As Boolean = (Interlocked.CompareExchange(_INTcacheExistence, 1, 1) = 1)
        Dim booIsNothing As Boolean
        Try
            SyncLock _OBJcache
                Dim colHandle As Collection = _OBJcache.colCache
            End SyncLock
        Catch
            booIsNothing = True
        End Try
        If booExistence AndAlso booIsNothing Then
            Return _ENUcacheExistence.beingCreated
        ElseIf booExistence AndAlso Not booIsNothing Then
            Return _ENUcacheExistence.available
        ElseIf Not booExistence AndAlso booIsNothing Then
            Return _ENUcacheExistence.notAvailable
        Else
            errorHandler_("Cache state is not valid", _
                          ClassName, _
                          "cacheExistence_", _
                          "Continuing: won't use the cache")
            Return _ENUcacheExistence.notAvailable
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Update cache
    '
    '
    Private Function cacheUpdate_(ByVal strFromString As String) As Boolean
        Select Case cacheExistence_()
            Case _ENUcacheExistence.available
            Case _ENUcacheExistence.beingCreated : Return False
            Case _ENUcacheExistence.notAvailable
                If Not cacheCreate_() Then Return False
        End Select
        SyncLock _OBJcache
            If _OBJcache.intCacheLimit = 0 Then Return False
        End SyncLock
        Dim objClone As qbVariableType
        Try
            objClone = Me.clone(True)
        Catch
            Return False
        End Try
        SyncLock _OBJcache
            With _OBJcache.colCache
                Dim intExcess As Integer = .Count - _OBJcache.intCacheLimit + 1
                Dim intIndex1 As Integer
                For intIndex1 = 1 To intExcess
                    .Remove(intIndex1)
                Next intIndex1
                Try
                    .Add(objClone, strFromString)
                Catch ex As Exception
                    errorHandler_("Cannot cache clone of object: " & _
                                Err.Number & " " & Err.Description, _
                                "cacheUpdate_", _
                                "Won't complete the cache operation")
                    Return False
                End Try
            End With
        End SyncLock
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' Verify that toString returns a fromString that creates the same info
    '
    '
    Private Function checkToFromString_(ByVal strToString As String) As Boolean
        Dim objQBvariableTypeTest As qbVariableType
        Try
            objQBvariableTypeTest = New qbVariableType
        Catch
            Return (False)
        End Try
        If Not objQBvariableTypeTest.fromString(strToString) Then
            errorHandler_("Cannot convert the following string back " & _
                            "to a test object using fromString: " & _
                            _OBJutilities.enquote(strToString), _
                            "toString", _
                            "Marking object not usable")
            Me.mkUnusable() : Return (False)
        End If
        Dim strExplanation As String
        If Not Me.compareTo(objQBvariableTypeTest, strExplanation) Then
            errorHandler_("toString output does not create a clone, " & _
                            "when used as fromString: " & _
                            "the toString output is " & _
                            _OBJutilities.enquote(strToString) & ", " & _
                            "while the toString for the test is " & _
                            _OBJutilities.enquote _
                            (objQBvariableTypeTest.ToString) & ": " & _
                            strExplanation, _
                            "toString", _
                            "Marking object not usable")
            Me.mkUnusable() : Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Check usability
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String) As Boolean
        If Not Me.Usable Then
            errorHandler_("Object is not usable", _
                          strProcedure, _
                          "A software error or object abuse has " & _
                          "occured, and this property or method cannot " & _
                          "proceed.")
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Clear the UDT member collection
    '
    '
    Private Function clearUDTmembers_() As Boolean
        Dim colHandle As Collection = CType(USRstate.objVarType, Collection)
        Dim intIndex1 As Integer
        With colHandle
            For intIndex1 = 1 To .Count
                If Not CType(CType(CType(.Item(intIndex1), _
                                         Collection), _
                                   Collection).Item(3), _
                             qbVariableType).disposeInspect Then
                    Return (False)
                End If
            Next intIndex1
        End With
        If Not OBJcollectionUtilities.collectionClear(CType(USRstate.objVarType, Collection)) Then
            Return (False)
        End If
    End Function

    ' -----------------------------------------------------------------------
    ' Initialize the bound list
    '
    '
    Private Function createBoundList_() As Boolean
        Try
            USRstate.colBounds = New Collection
        Catch
            errorHandler_("Cannot create the bounds collection: " & _
                            Err.Number & " " & Err.Description, _
                            "New_", _
                            "Object is not usable")
            Return (False)
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------------
    ' Interface to error handler
    '
    '
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String)
        errorHandler_(strMessage, Me.Name, strProcedure, strHelp)
    End Sub
    Private Overloads Shared Sub errorHandler_(ByVal strMessage As String, _
                                               ByVal strObjectName As String, _
                                               ByVal strProcedure As String, _
                                               ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, strObjectName, strProcedure, strHelp)
    End Sub

    ' -----------------------------------------------------------------------
    ' Extend the bound list
    '
    '
    Private Function extendBoundList_(ByVal intLBound As Integer, _
                                      ByVal intUBound As Integer) As Boolean
        With USRstate
            Try
                Dim colEntry As Collection
                colEntry = New Collection
                colEntry.Add(intLBound) : colEntry.Add(intUBound)
                .colBounds.Add(colEntry)
            Catch
                errorHandler_("Cannot extend the bound list for " & _
                              _OBJutilities.enquote(Me.Name) & ": " & _
                              Err.Number & " " & Err.Description, _
                              "extendBoundList_", _
                              "Marking object not usable")
                Me.mkUnusable() : Return (False)
            End Try
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Internal inspection
    '
    '
    Private Function inspection_(ByVal booBasic As Boolean) As Boolean
        Dim strReport As String
        If Me.inspect(strReport, booBasic:=booBasic) Then Return (True)
        errorHandler_("Internal inspection has failed: report follows" & _
                      vbNewLine & vbNewLine & _
                      strReport, _
                      "", _
                      "Marking object unusable")
        Me.mkUnusable() : Return (False)
    End Function

    ' -----------------------------------------------------------------------
    ' Tell caller if type is simple (scalar, unknown or null)
    '
    '
    Private Function isSimpleType_(ByVal enuType As ENUvarType) As Boolean
        If isScalarType(enuType) Then Return (True)
        If enuType = ENUvarType.vtUnknown Then Return (True)
        If enuType = ENUvarType.vtNull Then Return (True)
        Return (False)
    End Function

    ' -----------------------------------------------------------------------
    ' Tell caller whether string identifies a type
    '
    '
    Private Function isType_(ByVal strInstring As String) As Boolean
        Return (_OBJutilities.findWord(VARIABLETYPES, _
                                      vtPrefixRemove(strInstring), _
                                      booIgnoreCase:=True) _
               <> _
               0)
    End Function

    ' -----------------------------------------------------------------------
    ' Identifies our numeric .Net types
    '
    '
    Private Shared Function numericNetTypes_() As String
        Return (netType2Name("Boolean") & " " & _
               netType2Name("Byte") & " " & _
               netType2Name("Short") & " " & _
               netType2Name("Integer") & " " & _
               netType2Name("Single") & " " & _
               netType2Name("Double"))
    End Function

    ' -----------------------------------------------------------------------
    ' Random type
    '
    '
    Private Overloads Shared Function randomType_(ByVal booMayReturnVariant As Boolean, _
                                                  ByVal strTypes As String) _
            As ENUvarType
        With _OBJutilities
            Dim intIndex1 As Integer = .findWord(UCase(strTypes), "VARIANT")
            Dim strVariableTypes As String = _
                .phrase(strTypes, 1, intIndex1 - 1) & " " & _
                .phrase(strTypes, intIndex1 + 1) & " " & _
                CStr(IIf(booMayReturnVariant, "Variant", ""))
            Select Case vtPrefixRemove(UCase(.word(strVariableTypes, _
                                                    Math.Max(1, CInt(Rnd() * .words(strVariableTypes))))))
                Case "BOOLEAN" : Return ENUvarType.vtBoolean
                Case "BYTE" : Return ENUvarType.vtByte
                Case "INTEGER" : Return ENUvarType.vtInteger
                Case "LONG" : Return ENUvarType.vtLong
                Case "SINGLE" : Return ENUvarType.vtSingle
                Case "DOUBLE" : Return ENUvarType.vtDouble
                Case "STRING" : Return ENUvarType.vtString
                Case "ARRAY" : Return ENUvarType.vtArray
                Case "NULL" : Return ENUvarType.vtNull
                Case "UNKNOWN" : Return ENUvarType.vtUnknown
                Case "VARIANT" : Return ENUvarType.vtVariant
                Case Else
                    errorHandler_("Unexpected case", _
                                  ClassName, _
                                  "randomType_", _
                                  "Returning unknown")
                    Return ENUvarType.vtUnknown
            End Select
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Reset to the Unknown type
    '
    '
    Private Function reset_() As Boolean
        With USRstate
            .enuVariableType = ENUvarType.vtUnknown
            If Not (.objVarType Is Nothing) Then
                If Me.isUDT Then
                    clearUDTmembers_()
                Else
                    CType(.objVarType, qbVariableType).dispose()
                End If
                .objVarType = Nothing
            End If
        End With
        Return (Me.createBoundList_)
    End Function

    ' ----------------------------------------------------------------------
    ' .Net type name to generic name
    '
    '
    Private Shared Function type2NetName_(ByVal strType As String) As String
        Dim strTypeWork As String = UCase(strType)
        If Mid(strTypeWork, 1, 7) <> "SYSTEM." Then
            strTypeWork = "SYSTEM." & strTypeWork
        End If
        Select Case strTypeWork
            Case "SYSTEM.BOOLEAN" : Return "Boolean"
            Case "SYSTEM.BYTE" : Return "Byte"
            Case "SYSTEM.SHORT" : Return "Short"
            Case "SYSTEM.INTEGER" : Return "Integer"
            Case "SYSTEM.INT16" : Return "Short"
            Case "SYSTEM.INT32" : Return "Integer"
            Case "SYSTEM.INT64" : Return "Long"
            Case "SYSTEM.LONG" : Return "Long"
            Case "SYSTEM.SINGLE" : Return "Single"
            Case "SYSTEM.DOUBLE" : Return "Double"
            Case "SYSTEM.STRING" : Return "String"
            Case "SYSTEM.OBJECT" : Return "Object"
            Case Else
                errorHandler_("Invalid precise .Net type name " & _OBJutilities.enquote(strType), _
                              ClassName, _
                              "type2NetName", _
                              "Returning null string")
                Return ("")
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Resolves the user data type index object into an integer
    '
    '
    Private Function udtIndex2Integer_(ByVal objIndex As Object, _
                                       ByVal intMemberCount As Integer) As Integer
        Dim intIndex1 As Integer
        Try
            intIndex1 = CInt(objIndex)
        Catch
            Try
                intIndex1 = CInt(CType(CType(USRstate.objVarType, _
                                             Collection).Item(CStr(objIndex)), _
                                       Collection).Item(1))
            Catch : End Try
        End Try
        If intIndex1 <= 0 OrElse intIndex1 > intMemberCount Then
            errorHandler_("Invalid index " & _
                          _OBJutilities.object2String(objIndex), _
                          "udtIndex2Integer_", _
                          "Returning zero")
            Return (0)
        End If
        Return (intIndex1)
    End Function

    End Class