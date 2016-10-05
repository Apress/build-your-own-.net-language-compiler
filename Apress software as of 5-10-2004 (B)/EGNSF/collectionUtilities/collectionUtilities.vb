Option Strict

Imports System.Threading

' **************************************************************************
' *                                                                        *
' * Utility methods for working with Collections                           *
' *                                                                        *
' *                                                                        *
' * This object exposes a number of utility methods for working with the   *
' * Collection object including collections that contain nested            *
' * collections. This object can serialize and deserialize collections,    *
' * it can compare them for identity, it can clone-copy collections and it *
' * can treat collections as sets and as indexed dictionaries.             *
' *                                                                        *
' * When declared with events this object can raise events showing its     *
' * progress through large-scale collections.                              *
' *                                                                        *
' * This class supports an innovative "macro" facility in the form of its  *
' * method collectionCalculus which can support a complex series of        *
' * operations.                                                            *
' *                                                                        *
' * To keep the state of this object at a minimum, all work is done with   *
' * collections passed as parameters. Visualize this rather thin object    *
' * as a membrane or disguise which when encasing a collection             *
' * significantly enhances its overall capabilities. In particular, this   *
' * object is useful in improving legacy Visual Basic code that uses       *
' * complex and nesting collections.                                       *
' *                                                                        *
' * Note that unlike its fraternal objects utilities and windowsUtilities, *
' * collectionUtilities has state in the form of events and an instance    *
' * Name property. Therefore, this class must be created before use, either*
' * by declaring the instance As New, or assigning the instance to a New   *
' * collectionUtilities object.  This is required, whether or not the      *
' * object instance is declared WithEvents.                                *
' *                                                                        *
' * The rest of this header block addresses the following topics.          *
' *                                                                        *
' *                                                                        *
' *      *  Properties, methods and events exposed by this object          *
' *      *  The collection calculus                                        *
' *      *  Multithreading considerations                                  *
' *      *  Change record                                                  *
' *      *  Issues                                                         *
' *                                                                        *
' *                                                                        *
' * PROPERTIES, METHODS AND EVENTS EXPOSED BY THIS OBJECT ---------------- *
' *                                                                        *
' * Properties are named starting with an upper case letter; methods       *
' * use camelCase and start with a lower case letter: events use camelCase *
' * and end with Event.                                                    * 
' *                                                                        *
' * For reliability and security, the object is at all times, upon the     *
' * completion of the object constructor, either Usable or not Usable.     *
' *                                                                        *
' * Upon completion of a successful New constructor, the object is         *
' * Usable until a serious error occurs, the mkUnusable method is run,     *
' * or the dispose method is executed.  Most procedures will then check    *
' * usability, and refuse to proceed when the object is unusable.          *
' *                                                                        *
' * All methods, which do not otherwise return a value, return when        *
' * called as functions a Boolean success flag.                            *
' *                                                                        *
' *                                                                        *
' *      *  About: this read-only property returns information, about the  *
' *         class                                                          *
' *                                                                        *
' *      *  ClassName: this read-only property returns the class name      *
' *         (collectionUtilities)                                          *
' *                                                                        *
' *      *  collection2Set: this method transforms a collection into a set,*
' *         where a set is a collection without any duplicate entries.     *
' *                                                                        *
' *         Note that duplication testing uses the CompareAllowsClone      *
' *         property.                                                      *
' *                                                                        *
' *      *  collection2Report: this method converts a collection to a      *
' *         row and column report                                          *
' *                                                                        *
' *      *  collection2SetEvent: this event is raised when each item is    *
' *         checked for eligibility for membership in the set: provides    *
' *         the object name, item number, the count of items at the cur-   *
' *         rent level and the level number. Useful for progress reporting.*
' *                                                                        *
' *         Note: since collection2Set works backward (from Count to 1 in  *
' *         "flat" collections without subcollections and in postorder in  *
' *         tree-structured collection with subcollections), the progress  *
' *         report in this event, unlike similar events exposed by this    *
' *         object, will count down, from Count to 1.                      * 
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collection2String: this method converts  a collection to a     *
' *         string                                                         *
' *                                                                        *
' *      *  collection2StringEvent: this event is raised when any          *
' *         string-serialized collection item has been appended to the     *
' *         collection during collection2String: provides the object Name, *
' *         the count of items at the current level, the current item      *
' *         index and the level number. Useful in progress reporting.      *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collection2XML(collection): this method converts  a collection *
' *         to an eXtended Markup Language (XML) tag.                      *
' *                                                                        *
' *         Each item is converted to a subtag with the name Itemnn where  *
' *         nn is the item number; the item number is expanded to the size *
' *         of the largest item number and zero-filled.                    *
' *                                                                        *
' *         Each item is in the format <Itemnn>value</Itemnn> where value  *
' *         is the string form of the item object, with decoration in      *
' *         effect.                                                        *
' *                                                                        *
' *         By default, each subcollection in the main collection is       *
' *         converted to a subtag in the same format as the main tag but   *
' *         with the name subcollectionnn, where nn is a sequence number.  *
' *                                                                        *
' *         The following overloads are supported:                         *
' *                                                                        *
' *         + collection2XML(collection) creates a tag with the overall    *
' *           name anonymousCollection. If the collection contains sub-    *
' *           collections they are fully expanded to XML tags.             *
' *                                                                        *
' *         + collection2XML(collection, name) creates a tag with the      *
' *           name supplied. If the collection contains subcollections they*
' *           are fully expanded to XML tags.                              *
' *                                                                        *
' *         + collection2XML(collection, name, False) creates a tag with   *
' *           the overall name supplied. If the collection contains sub-   *
' *           collections they are not fully expanded to XML tags: instead,*
' *           the subcollections are converted using object2String.        *
' *                                                                        *
' *         + collection2XML(collection, name, expansionOption, False)     *
' *           creates a tag with the overall name supplied. If the         *
' *           collection contains subcollections they are fully expanded to*
' *           XML tags if expansionOption is True. A newline is NOT placed *
' *           between elements and the tag returned is packed and not      *
' *           formatted as shown below.                                    *
' *         
' *         If the collection is Nothing, then the tag <name>Nothing</name>*
' *         is produced.                                                   *
' *                                                                        *
' *         Suppose a collection contains an integer, a string, and a sub- *
' *         collection containing one integer. collection2XML will produce *
' *         the following tag.                                             *
' *                                                                        *
' *              <anonymousCollection>                                     *
' *                   <item1>System.Int32(32767)</item1>                   *
' *                   <item2>System.String("Example")</item2>              *
' *                   <item3>                                              *
' *                       <subcollection1>                                 *
' *                           <item1>System.Int32(14)</item1>              *
' *                       </subcollection1>                                *
' *                   </item3>                                             *
' *              </anonymousCollection>                                    *
' *                                                                        *
' *      *  collectionAppend(collection1, collection2): appends all the    *
' *         members of collection 1 to the end of collection 2.            *
' *                                                                        *
' *      *  collectionAppendEvent: this event is raised after each item has*
' *         been added to the appended collection during collectionClear:  *
' *         provides the object Name, the number of items being appended   *
' *         at the current level and the index of the new item.  Useful for*
' *         progress reporting.                                            *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionCalculus(expression, collectionList): this method    *
' *         executes an expression in the "collection calculus" which      *
' *         performs a series of our collection operations.                *
' *                                                                        *
' *         This is a simple or nested expression in the overall form      *
' *         methodName(methodOperands) which yields, depending on the last *
' *         methodName executed, a collection, a string or a Boolean value.*
' *                                                                        *
' *         This method returns Nothing, a Collection or an object as the  *
' *         final result of its last calculus operation.                   *
' *                                                                        *
' *         See the section, below, on the collection calculus for more    *
' *         information.                                                   *
' *                                                                        *
' *      *  collectionCalculusEvent: this event is raised after each step  *
' *         in the evaluation of a collectionCalculus expression is        *
' *         complete. It is supplied with the object name, the operation   *
' *         performed in the step and the result of the step (as an        *
' *         object).                                                       *
' *                                                                        *
' *         The operation will be a method name, 0_PUSHLITERAL or          *
' *         1_PUSHBYREF. The result of the step will be the stack top      *
' *         after the operation or Nothing.                                *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionCalculusWithReport(expression, report,               *
' *         collectionList): this method, like collectionCalculus as above,*
' *         executes an expression in the "collection calculus" which      *
' *         performs a series of our collection operations. This method as *
' *         above supports the creation of a list of named collections.    *
' *                                                                        *
' *         In addition, the collectionCalculusWithReport method places    *
' *         a trace or report of the methods executed in the calculus in   *
' *         its report parameter. This parameter must be a System.String,  *
' *         however, it must be passed as an object and by reference. If   *
' *         your code has Option Strict in effect, you must use the CObj   *
' *         function to pass the report parameter as in CObj(strReport).   *
' *                                                                        *
' *         This method returns Nothing, a Collection or an object as the  *
' *         final result of its last calculus operation.                   *
' *                                                                        *
' *      *  collectionClear: this method clears collections including      *
' *         nested objects                                                 *
' *                                                                        *
' *      *  collectionClearEvent: this event is raised after each item has *
' *         been removed from the collection during collectionClear:       *
' *         provides the object Name, the number of items being removed    *
' *         at the current level, the index of the removed item and the    *
' *         nesting level. Useful for progress reporting.                  *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionCompare: this method compares collections for        *
' *         identity item-by-item, using the CompareAllowsClone property.  *
' *                                                                        *
' *         When CompareAllowsClone is True then scalars are compared using*
' *         their string value. Collections are compared by recursive use  *
' *         of this method. Objects other than collections always cause    *
' *         this method to return False when CompareAllowsClone is True.   *
' *                                                                        *
' *         When CompareAllowsClone is False then scalars are compared as  *
' *         
' *      *  collectionCompareEvent: this event is raised when each pair of *
' *         collection items has been compared: provides the object Name,  *
' *         the number of items at the current level of comparision,       *
' *         the index at which the comparision occurs and the nesting      *
' *         level. Useful for progress reports.                            *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionComplement: this method forms the complement of  a   *
' *         collection A with respect to a collection B, returning the     *
' *         complement as a function value. The complement is the set of   *
' *         all members of B which are not in A.                           *
' *                                                                        *
' *         For the purpose of forming the intersection, the               *
' *         CompareAllowsClone property determines whether objects other   *
' *         than collections are compared for IS identity.                 *
' *                                                                        *
' *         The result set never contains duplicate entries, even when     *
' *         the input sets contain duplicates.                             *
' *                                                                        *
' *         Note that the complement algorithm is slow for large           *
' *         collections at this writing.                                   *
' *                                                                        *
' *      *  collectionComplementEvent: this event is raised during the     *
' *         creation of set complements: it provides the object Name, the  *
' *         number of items at the current level, the index of the         *
' *         currently examined item and the nesting level. Useful for      *
' *         progress reports.                                              *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionCopy: this method can copy many collections.         *
' *                                                                        *
' *         When applied to a source collection, this method creates and   *
' *         returns a comprehensive "deep clone" of the collection. This is*
' *         a totally new collection such that:                            *
' *                                                                        *
' *         + Any subcollections are cloned recursively and not just       *
' *           referenced.                                                  *
' *                                                                        *
' *         + All scalar entries (defined as entries that convert without  *
' *           error to a string) are copied by value                       *
' *                                                                        *
' *         + For any other type of object, collectionCopy will clone the  *
' *           object as long as the object exposes a clone method with     *
' *           a return value that is an object of some type.               *
' *                                                                        *
' *      *  collectionCopyEvent: this event is raised when each item has   *
' *         been copied: it provides the object Name, the count of items at*
' *         the current level, the current item index and the level number.*
' *         Useful in progress reporting.                                  *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionDepth: this method returns the collection depth as   *
' *         its maximum number of levels                                   *
' *                                                                        *
' *      *  collectionDepthEvent: this event is raised during the determin-*
' *         ation of the maximum collection depth: it provides the object  *
' *         Name, the number of items at the current level, the index of   *
' *         the current item and the nesting level. Useful for progress    *
' *         reports.                                                       *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionFind: this method searches an argument collection    *
' *         for a member in two different modes: "clone OK" and "IS mode": *
' *                                                                        *
' *              + The default "clone OK mode" searches for Visual Basic   *
' *                scalars of type Boolean, Byte, Short, Integer, Long,    *
' *                Single, Double or String, or VB collections.            *
' *                                                                        *
' *                In this mode, scalars are reported as found when they   *
' *                match items after both are converted to strings.        *
' *                
' *              + In the "IS" mode, members of the compared collections   *
' *                can have any type. Value (scalar) objects are compared  *
' *                as above, but reference objects (including collections) *
' *                are compared for identity using the IS operator. In this*
' *                mode, a reference object search will report success     *
' *                only when the reference object is actually in the       *
' *                searched collection.                                    *
' *                                                                        *
' *         When the scalar or collection is found at the top level of the *
' *         argument collection, this method returns its Integer index.    *
' *                                                                        *
' *         When the scalar or collection is found inside a subcollection, *
' *         this method returns its String "Dewey Decimal" number. This is *
' *         a series of numbers separated by periods indicating the level  *
' *         at which the item was found and its index at each level.       *
' *                                                                        *
' *         For example, in the collection ("Item 1", (2, 4)), "Item 1"'s  *
' *         collectionFind value will be the integer 1 while the find value*
' *         for 4 will be the string "2.2".                                *
' *
' *      *  collectionFindEvent: this event is raised during the           *
' *         collectionFind search: it provides the object Name, the        *
' *         number of items at the current level, the index of             *
' *         the current item being checked and the nesting level. Useful   *
' *         for progress reports.                                          *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionIntersection: this method forms the intersection of a*
' *         collection A with respect to a collection B returning the      *
' *         intersection as a function value. The intersection is the set  *
' *         of all members in both A and B.                                *
' *                                                                        *
' *         For the purpose of forming the intersection, the               *
' *         CompareAllowsClone property determines whether objects other   *
' *         than collections are compared for IS identity.                 *
' *                                                                        *
' *         The result set never contains duplicate entries, even when     *
' *         the input sets contain duplicates.                             *
' *                                                                        *
' *         Note that the intersection algorithm is slow for large         *
' *         collections at this writing.                                   *
' *                                                                        *
' *      *  collectionIntersectionEvent: this event is raised during the   *
' *         creation of set intersections: it provides the object Name, the*
' *         number of items at the current level, the index of the         *
' *         currently examined item and the nesting level. Useful for      *
' *         progress reports.                                              *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionItemSet: this method modifies one item in a          *
' *         collection or contained subcollection, which may be identified *
' *         by an integer, or, to identify members in contained subcollec- *
' *         tions, may be identified by a Dewey Decimal number.            *
' *                                                                        *
' *         The modified collection is returned on success; Nothing is     *
' *         returned to indicate an error.                                 *
' *                                                                        *
' *      *  collectionItemSetEvent: this event is raised during the        *
' *         modification of an item by the collectionItemSet method. It    *
' *         is provided the object name, the Dewey Decimal position of the *
' *         change, the old value and the new value.                       *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionParseEvent: this event is raised during the parse    *
' *         of collection calculus expressions on each attempt to parse    *
' *         a grammar category, and it identifies the object name, the goal*
' *         symbol and the current token index.                            *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionTypes: this method creates the set of distinct types *
' *         found in its argument collection as a new collection that      *
' *         contains the system's type name in a string in 0..n items.     * 
' *                                                                        *
' *      *  collectionTypeEvent: this event is raised during the search for*
' *         distinct item types in collectionType: it provides the object  *
' *         Name, the number of items at the current level, the index of   *
' *         the current item and the nesting level. Useful for progress    *
' *         reports.                                                       *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  collectionUnion: this method forms the union of a collection A *
' *         with respect to a collection B returning the union as a        *
' *         function value. The union is the set of all members in both A  *
' *         and B.                                                         *
' *                                                                        *
' *         For the purpose of forming the union, the CompareAllowsClone   *
' *         property determines whether objects other than collections are *
' *         compared for IS identity.                                      *
' *                                                                        *
' *         The result set never contains duplicate entries, even when     *
' *         the input sets contain duplicates.                             *
' *                                                                        *
' *         Note that the union algorithm is slow for large collections at *
' *         this writing.                                                  *
' *                                                                        *
' *      *  collectionUnionEvent: this event is raised during the creation *
' *         of set unions: it provides the object Name, the number of items*
' *         at the current level, the index of the currently examined item *
' *         and the nesting level. Useful for progress reports.            *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  CompareAllowsClone: this read-write property defines the       *
' *         comparision rule used by the following methods in comparing    *
' *         collection members: collectionCompare, collectionFind,         *
' *         collectionComplement, collectionIntersection and               *
' *         collectionUnion.                                               *
' *                                                                        *
' *         When this property is True the following comparision rules are *
' *         in effect.                                                     *
' *                                                                        *
' *              + Scalars of type Boolean, Byte, Short, Integer, Single,  *
' *                Double, and String don't have to have identical type    *
' *                when compared. Instead, the compared scalars are        *
' *                converted to strings and the strings must match exactly.*
' *                                                                        *
' *              + Collection objects are considered identical when the    *
' *                collectionCompare method indicates that their members   *
' *                are the same. Since collectionCompare uses the True     *
' *                value of the CompareAllowsClone property, collection    *
' *                objects, here, will be considered identical when they   *
' *                contain scalars and subcollections only, and each       *
' *                subcollection is a clone of the corresponding member.   *
' *                                                                        *
' *              + Two objects that aren't both collections are considered *
' *                nonidentical.                                           *
' *                                                                        *
' *         When this property is False the following comparision rules are*
' *         in effect.                                                     *
' *                                                                        *
' *              + Scalars of type Boolean, Byte, Short, Integer, Single,  *
' *                Double, and String are compared as above.               *
' *                                                                        *
' *              + Reference objects including collections must be         *
' *                instance-identical (using A Is B).                      *
' *                                                                        *
' *         The CompareAllowsClone property defaults to True.              *
' *                                                                        *
' *      *  dewey2Member: this method returns the member of the collection *
' *         (or, one of its subcollections) accessed by an integer or a    *
' *         Dewey Decimal number using the syntax dewey2Member(i) or       *
' *         dewey2Member(dewey)                                            *
' *                                                                        *
' *         When passed an integer (or a string that converts to an integ- *
' *         er) this method will return the item at the indicated          *
' *         position, at the top or only collection level.                 *
' *                                                                        *
' *         When passed a Dewey Decimal number as a series of indexes      *
' *         separated by periods this method will return subcollection     *
' *         members. For example, when Item(1) of a collection is a sub-   *
' *         collection, and this subcollection contains a sub-sub-         *
' *         collection at Item(2), the Dewey Decimal number 1.2.3 will re- *
' *         turn the third element of the sub-sub-collection.              *
' *                                                                        *
' *         The overloaded syntax dewey2Member(dewey, parent, index) can   *
' *         return the parent collection of the found member and the       *
' *         item number of the found member.                               *
' *                                                                        *
' *      *  dispose: this method marks the object not usable. Normally,    *
' *         methods with this name are responsible for disposing reference *
' *         objects in state. However, now, and in the foreseeable future, *
' *         the thin state of this object (consisting only of the usability*
' *         flag and the name) contains no reference objects, and this     *
' *         method should be used only to conform to a general expectation.*
' *                                                                        *
' *      *  extendIndex: this method extends a collection structured as a  *
' *         dictionary according to the rules of the isIndex method: these *
' *         collections consist of two-item subcollections such that       *
' *         item(1) is a key and item(2) is data.                          *
' *                                                                        *
' *         This method has the syntax extendIndex(key,data). The key      *
' *         may be any Visual Basic object that converts without error to  *
' *         a string, and it is converted to a string both for the key, and*
' *         for storage in its subcollection. If it is a number, it is     *
' *         prefixed with an asterisk.                                     *
' *                                                                        *
' *         The data field may be any object.                              *
' *                                                                        *
' *      *  isIndex: this method returns True when the collection is an    *
' *         indexed dictionary such that:                                  *
' *                                                                        *
' *         + It contains at least one entry                               *
' *                                                                        *
' *         + Each entry is a subcollection, containing two items such that*
' *                                                                        *
' *           - Item(1) is a string which acts when tested as the item key *
' *           - Item(2) is any object                                      *
' *                                                                        *
' *         The basic syntax is isIndex(collection). When called with the  *
' *         overloaded syntax isIndex(collection, explanation), the        *
' *         explanation parameter, which should be a string, is set to an  *
' *         explanation of why the collection is not an indexed dictionary,*
' *         or a null string.                                              *
' *                                                                        *
' *      *  list2Collection: this method creates a collection from a list  *
' *         of objects of any type.                                        *
' *                                                                        *
' *         Note that where it is tedious and repetitious to (properly)    *
' *         create collections in error handlers, this function will       *
' *         automate this; list2Collection() without operands creates a    *
' *         collection in an error handler, returning Nothing when an      *
' *         error occurs.                                                  *
' *                                                                        *
' *      *  mkRandomCollection: this method creates a random collection.   *
' *         It exposes the following overloads:                            *
' *                                                                        *
' *         + mkRandomCollection: creates a random collection, with from   *
' *           0 to 100 items. Each item can be a Boolean, Byte, Short,     *
' *           Integer, Long, Single, Double, String, Collection or         *
' *           collectionUtilities object. When an item is a collection,    *
' *           nesting is limited to 8 levels.                              *
' *                                                                        *
' *         + mkRandomCollection(minItems, maxItems): creates a random     *
' *           collection as above; but with from minItems to maxItems.     *
' *                                                                        *
' *         + mkRandomCollection(minItems, maxItems, types): creates a     *
' *           random collection as above; but restricted to the space-sepa-*
' *           rated list of types, which can contain any combination of the*
' *           types Boolean, Byte, Short, Integer, Long, Single, Double,   *
' *           String, Collection, or collectionUtilities.                  *
' *                                                                        *
' *         + mkRandomCollection(minItems, maxItems, types, nesting):      *
' *           creates a random collection as above; but when collections   *
' *           are nested, the nesting is limited to the value specified.   *
' *           This value may be -1 for no limit except as randomly         *
' *           generated. Note that mkRandomCollection(minItems, maxItems,  *
' *           "Collection", -1) will cause a loop.                         *
' *                                                                        *
' *      *  mkRandomCollectionEvent: this event is raised after each       *
' *         item is added to the random collection: it provides the        *
' *         object name, the randomly-generated item count at the current  *
' *         level, the current item index, and the nesting level.          *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  mkUnusable: this method makes the object not usable.           *
' *                                                                        *
' *      *  Name: this read-write property identifies the instance of the  *
' *         collectionUtilities. The object name defaults to               *
' *         collectionUtilitiesnnnn date time where nnnn is a sequence     *
' *         number.                                                        *
' *                                                                        *
' *      *  retrieveFromIndex(c,k): when c is a collection structured      *
' *         according to the rules of the isIndex method, this method      *
' *         returns the value of the dictionary for key k. If k is a number*
' *         it is prefixed with an asterisk to retrieve the value.         *
' *                                                                        *
' *      *  scalar2Collection(object): this method creates a new collection*
' *         containing one item set to the object.                         *
' *                                                                        *
' *      *  string2Collection: this method converts a string in the a      *
' *         "serialized" form that is produced by collection2String back to*
' *         a collection.                                                  *
' *                                                                        *
' *      *  string2CollectionEvent: this event is raised when each item in *
' *         the input string has been converted; it provides the item count*
' *         at the current level, the current item index, and the nesting  *
' *         level. Useful for progress reports.                            *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  test(report): this method executes a stress test on the        *
' *         class and returns True on success or False on failure. A test  *
' *         report is returned in the Byref string, report.                * 
' *                                                                        *
' *         By default, 100 stress tests are executed in such a way that   *
' *         identical results always occur. The optional overload          *
' *         syntax test(report,count) may specify a different count.       *
' *         Use the Randomize statement in the calling environment for     *
' *         unpredictable results.                                         *
' *                                                                        *
' *         Note that this method runs the test using the object instance, *
' *         and, it changes the state of the object instance.              *
' *                                                                        *
' *      *  testEvent: this event is raised on each test; it provides the  *
' *         item count at the current level, and the current item index.   *
' *         Useful for progress reports.                                   *
' *                                                                        *
' *         Declare this object WithEvents to get this event.              *
' *                                                                        *
' *      *  TestExpression: this shared and read-only property returns the *
' *         expression used by the test method.                            *
' *                                                                        *
' *      *  Usable: this read-only property identifies the usability of    *
' *         the object, as described above.                                *
' *                                                                        *
' *                                                                        *
' * THE COLLECTION CALCULUS ---------------------------------------------- *
' *                                                                        *
' * The collectionCalculus method executes a series of operations on       *
' * collections which can accomplish complex objectives in a simple and    *
' * transparent manner. It may identify one or more named collections,     *
' * referenced in the collection calculus.                                 *
' *                                                                        *
' * The collectionCalculusWithReport method performs the functions of the  *
' * collectionCalculus method and also creates a report of the operations  *
' * performed. You can use your own calculus expressions to tag this       *
' * report with extra material.                                            *
' *                                                                        *
' * This method has the syntax collectionCalculus(expression) or the       *
' * extended syntaxes, collectionCalculus(expression, list) or             *
' * collectionCalculusWithReport(expression, report, list).                *
' *                                                                        *
' * The expression should be the name of one of the following operations   *
' * which as you can see correspond to a subset of the methods of this     *
' * object, a compound expression where the method's output is "piped"     *
' * to another method, or a series of operations that are separated by     *
' * colons.                                                                *
' *                                                                        *
' * The list may consist of 2..10 parameters (always supplied as an even   *
' * number of parameters) in the form collectionName1,collection1,         *
' * collectionName2,collection2...collectionNamem,collectionm. Each list   *
' * pair specifies the string name of a collection and its value, and the  *
' * string names can be used in collection calculus expressions anywhere a *
' * collection is valid. In addition (and unlike collections produced as   *
' * results) these collections are used "by reference" and may be changed. *
' *                                                                        *
' * Each string name must be a Visual Basic identifier that starts with an *
' * underscore such as _Collection1.                                       *
' *                                                                        *
' * The extended syntax collectionCalculusWithReport(expression, report,   *
' * list) is available. It produces a trace or report of the calculus      *
' * operations and their results in its report parameter, which must be    *
' * a System.String, and which is received by reference.                   *
' *                                                                        *
' * Note that the size of the by-reference collection list as supplied to  *
' * collectionCalculus is limited to five named collections. An advantage  *
' * of the collectionCalculusWithReport method is that it does not limit   *
' * the size of the by-reference list.                                     *
' *                                                                        *
' * Note that for each operation that has a collection as an argument, we  *
' * identify the argument as passable by value, by reference, or in either *
' * of these modes:                                                        *
' *                                                                        *
' *                                                                        *
' *      *  If the argument is an inner collection operation which makes   *
' *         a collection, the argument was passed by value.                *
' *                                                                        *
' *      *  If the argument is the name of a collection, the argument was  *
' *         passed by reference.                                           *
' *                                                                        *
' *                                                                        *
' * Consistently, an operation that modifies a collection will push the    *
' * result on a stack when its argument is by value, making the result     *
' * available to nesting functions, as in                                  *
' * collectionUnion(mkRandomCollection, mkRandomCollection); it will modify*
' * the named collection when its argument is by reference.                *
' *                                                                        *
' * Here are the operations supported in the collectionCalculus.           *
' *                                                                        *
' *                                                                        *
' *      *  collection2Report(collection): converts the collection to a    *
' *         columnar report. The collection may be passed by reference or  *
' *         by value. The report is stacked as a string.                   *
' *                                                                        *
' *      *  collection2Set(collection): removes duplicates from the        *
' *         collection. The collection may be passed by reference or by    *
' *         value. The result is stacked when the collection was passed by *
' *         value; it modifies the named collection when the collection was*
' *         passed by reference.                                           *
' *                                                                        *
' *      *  collection2String(collection): returns the serialized          *
' *         collection where the collection argument is by value or by     *
' *         reference.                                                     *
' *                                                                        *
' *         The result is always stacked.                                  *
' *                                                                        *
' *         Note that the string produced is in the "unreadable" format    *
' *         created by the collection2String method with booReadable:=False*
' *         in effect; see also collection2StringReadable.                 *
' *                                                                        *
' *      *  collection2StringReadable(collection): returns the readable,   *
' *         serialized collection where the collection argument is by      *
' *         value or by reference.                                         *
' *                                                                        *
' *         The result is always stacked.                                  *
' *                                                                        *
' *         Note that the string produced is in the "readable" format      *
' *         created by the collection2String method with booReadable:=True *
' *         in effect; see also collection2String.                         *
' *                                                                        *
' *      *  collectionAppend(collection1, collection2): appends all the    *
' *         members of collection2 to the end of collection1.              *
' *                                                                        *
' *         Collection1 and collection2 may be passed by reference as the  *
' *         names of collections or by value as nested functions.          *
' *                                                                        *
' *         If collection1 is passed by reference using a name, it is      *
' *         modified, and returned as the value of this method. If         *
' *         collection1 is passed by value it is unchanged, and the        *
' *         append result is returned.                                     *
' *                                                                        *
' *      *  collectionClear(collection): clears the collection. The        *
' *         collection must be passed by reference and it is modified;     *
' *         the stack is not affected.                                     *
' *                                                                        *
' *      *  collectionCompare(collection1,collection2): returns True when  *
' *         collection1 is itemwise identical to collection2, False        *
' *         otherwise.  Both collections may be passed by value or by      *
' *         reference. The result of True or False is stacked.             *
' *                                                                        *
' *      *  collectionComplement(collection1,collection2): forms and       *
' *         returns (stacks) the complement of collection1 with respect to *
' *         collection2; this is the set of all collection2 members not    *
' *         found in collection1.                                          *
' *                                                                        *
' *      *  collectionCopy(collection): clones the collection where the    *
' *         collection argument is by value or by reference.               *
' *                                                                        *
' *         The syntax collectionCopy(collection,name) will add or replace *
' *         the new collection in the method list of named collections.    *
' *         It will not alter the stack.                                   *
' *                                                                        *
' *         The syntax collectionCopy(collection) will stack a copy of the *
' *         named collection or value collection.                          *
' *                                                                        *
' *         Note the difference between collectionSave and this operation; *
' *         this operation makes a duplicate collection whereas            *
' *         collectionSave will merely set the name to reference the       *
' *         collection.                                                    *
' *                                                                        *
' *      *  collectionDepth(collection): stacks the maximum nesting levels *
' *         in the collection, passed by value or by reference.            *
' *                                                                        *
' *      *  collectionFind(collection,object): searches the collection for *
' *         a scalar or collection and returns the integer toplevel index  *
' *         or Dewey Decimal number. If the target object cannot be found  *
' *         this method returns the Integer value 0.                       *
' *                                                                        *
' *         The return result is stacked.                                  *
' *                                                                        *
' *         The collection can be identified by reference or by value.     *
' *                                                                        *
' *      *  collectionIntersection(collection1,collection2): forms and     *
' *         returns (stacks) the intersection of collection1 and           *
' *         collection2; this is the set of all members in both            *
' *         collections.                                                   *
' *                                                                        *
' *      *  collectionItemSet(collection, deweyNumber, value): changes the *
' *         item, in the collection, identified by deweyNumber, to contain *
' *         value. The collection may be identified by name or obtained    *
' *         from a nested operation. The modified collection is returned on*
' *         the stack.                                                     * 
' *                                                                        *
' *      *  collectionSave(collection, name): saves the collection (without*
' *         making a clone) into the ByRef collection that is identified by*
' *         the name.                                                      *
' *                                                                        *
' *         Note the difference between collectionCopy and this operation; *
' *         collectionCopy makes a duplicate collection whereas this       *
' *         operation will merely set the name to reference the collection.*
' *                                                                        *
' *      *  collectionTypes(collection): returns the new collection of the *
' *         distinct types found in collection, which is passed by value or*
' *         by reference: in this syntax the type collection is stacked.   *
' *                                                                        *
' *         The syntax collectionTypes(collection,name) will add or replace*
' *         the new collection in the method list of named collections.    *
' *         This syntax won't change the stack.                            *
' *                                                                        *
' *      *  collectionUnion(collection1,collection2): forms and returns    *
' *         (stacks) the union of collection1 and collection2; this is the *
' *         set of all members in either collection.                       *
' *                                                                        *
' *      *  dewey2Member(collection,dewey): retrieves the member of the    *
' *         byRef or byVal collection identified by the Dewey Decimal      *
' *         number.                                                        *
' *                                                                        *
' *      *  dumpByRef: this method displays the contents of the named      *
' *         (ByRef) collections: note that it does not change the stack.   *
' *                                                                        *
' *      *  extendIndex(c,k,o): this method treats the collection c as a   *
' *         dictionary using the rules of the isIndex method. It adds one  *
' *         keyed entry, with k being the key, as a two-item subcollection *
' *         containing the key and the object specified in o.              *
' *                                                                        *
' *      *  isIndex(c): stacks True if the collection c is structured as a *
' *         dictionary using the rules of the isIndex method exposed by the*
' *         collection object: in these rules, c must consist of one or    *
' *         more two-item subcollections, where Item(1) is a key and       *
' *         Item(2) is data. If c is anything else, this calculus method   *
' *         returns False.                                                 *
' *                                                                        *
' *      *  list2Collection: creates a collection from a list of 0..10     *
' *         scalar values of the types Boolean, Byte, Short, Integer, Long,*
' *         Single, Double, or String.                                     *
' *                                                                        *
' *         Note: names of ByRef collections in this list won't be expanded*
' *         into collections. Instead they will be appended to the list as *
' *         strings.                                                       *
' *                                                                        *
' *      *  list2Collection: creates a collection from a list of 0..10     *
' *         scalar values of the types Boolean, Byte, Short, Integer, Long,*
' *         Single, Double, or String.                                     *
' *                                                                        *
' *         Note: names of ByRef collections in this list won't be expanded*
' *         into collections. Instead they will be appended to the list as *
' *         strings.                                                       *
' *                                                                        *
' *      *  mkRandomCollection(minItems,maxItems,types,nesting): creates   *
' *         a new, random collection:                                      *
' *                                                                        *
' *         + minItems should contain the minimum item count               *
' *                                                                        *
' *         + maxItems should contain the maximum item count               *
' *                                                                        *
' *         + types should contain the types permitted as a space-separated*
' *           list of the allowable types, selected from Boolean, Byte,    *
' *           Short, Integer, Long, Single, Double, String, Collection, or *
' *           collectionUtilities.                                         *
' *                                                                        *
' *         + nesting should contain the maximum level at which sub-       *
' *           collections may nest; nesting may be -1 for no limit.        *
' *                                                                        *
' *         The extended syntax mkRandomCollection(minItems,maxItems,types,*
' *         nesting,name) will insert the random collection in the method  *
' *         list under the identified name.                                *
' *                                                                        *
' *      *  rem: displays the current value in the stack, the product of   *
' *         the preceding operations. If the value is a Visual Basic       *
' *         scalar it is displayed as-is. If the value is a collection,    *
' *         it is displayed using collection2String with readability. All  *
' *         other values are displayed as type(value).                     *
' *                                                                        *
' *      *  rem(string): places the string in the report when              *
' *         collectionCalculusWithReport is used as a additional comments. *
' *                                                                        *
' *      *  string2Collection(string): stacks the new collection obtained  *
' *         by deserializing the string to a collection.                   *
' *                                                                        *
' *         The string should be in the format produced by                 *
' *         collection2String.                                             *
' *                                                                        *
' *                                                                        *
' * MULTITHREADING CONSIDERATIONS ---------------------------------------- *
' *                                                                        *
' * This object has state in the form of an instance name and events.      *
' * Multiple instances of this object may run in parallel threads; one     *
' * instance may not run in more than one thread.                          *
' *                                                                        *
' *                                                                        *
' * C H A N G E   R E C O R D -------------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE                        *
' * --------   ----------     -------------------------------------------- *
' * 07 05 03   Nilges         1.  Created this standalone class            *
' *                                                                        *
' * 07 16 03   Nilges         1.  Backed up this edition to DVD and also to*
' *                               C:\EGNSF\COLLECTIONUTILITIES\ARCHIVE\    *
' *                               07162003\version1                        *
' *                                                                        *
' *                           2.  Changed collection2String                *
' *                                                                        *
' *                           3.  Added dewey2Member                       *
' *                                                                        *
' *                           4.  Added collectionItemSet                  *
' *                                                                        *
' *                           5.  Changed dewey2Member                     *
' *                                                                        *
' * 07 23 03   Nilges         1.  Modified format of string display in     *
' *                               object2XML                               *
' *                                                                        *
' * 07 26 03   Nilges         1.  Backed up this edition to CD and also to *
' *                               C:\EGNSF\COLLECTIONUTILITIES\ARCHIVE\    *
' *                               07262003\version2                        *
' *                                                                        *
' *                           2.  Added isIndex                            *
' *                                                                        *
' *                           3.  Added extendIndex                        *
' *                                                                        *
' *                           4.  Added retrieveFromIndex                  *
' *                                                                        *
' * 09 29 03   Nilges         1.  Changed collection2String                *
' *                                                                        *
' * 10 18 03   Nilges         1.  An OrElse in string2Collection was in the*
' *                               wrong order such that a Nothing was      *
' *                               evaluated                                *
' *                                                                        *
' * 01 23 04   Nilges         1.  Added collection2XML                     *
' *                                                                        *
' * 02 08 04   Nilges         1.  In copying collections attempt to use    *
' *                               a clone method on unsupported item types *
' *                                                                        *
' * 05 04 04   Nilges         1.  Changed collection2Report                *
' *                                                                        *
' * I S S U E S ---------------------------------------------------------- *
' *   DATE         MODULE        POSTER     DESCRIPTION                    *
' * --------   -------------- ------------  ------------------------------ *
' * 07 15 03                  Nilges        Collection methods are slow    *
' *                                         for large collections.         *
' *                                                                        *
' * 07 16 03                  Nilges        This class has not been tested *
' *                                         completely up to the stress    *
' *                                         test. All the "collection      *
' *                                         calculus" tests supplied in    *
' *                                         the collectionTest utility     *
' *                                         work.                          * 
' *                                                                        *
' * 07 16 03                  Nilges        compareAllowsClone has not been*
' *                                         tested.                        *
' *                                                                        *
' * 07 16 03                  Nilges        Consider adding quoted strings *
' *                                                                        *
' * 07 16 03                  Nilges        Consider adding an interpret   *
' *                                         operation                      *
' **************************************************************************

Public Class collectionUtilities

#Region " General declarations "

    ' ***** Shared *****
    Private Shared _INTsequence As Integer
    Private Shared _OBJutilities As utilities.utilities
    
    ' ***** State *****
    Private Structure TYPstate
        Dim booUsable As Boolean
        Dim strName As String
        Dim booCompareAllowsClone As Boolean
    End Structure 
    Private USRstate As TYPstate   
    
    ' ***** Constants *****
    ' --- Easter Egg
    Private Const CLASS_NAME As String = "collectionUtilities"
    Private Const ABOUT_INFO As String = _
    "This object exposes a number of utility methods for working with the " & _
    "Collection object including collections that contain nested " & _
    "collections. This object can serialize and deserialize collections, " & _
    "it can compare them for identity, it can clone-copy collections and " & _
    "it can treat collections as sets and as indexed dictionaries." & _
    vbNewline & vbNewline & _
    "When declared with events this object can raise events showing its " & _
    "progress through large-scale collections" & _
    vbNewline & vbNewline & _
    "This class supports an innovative ""macro"" facility in the form of its " & _
    "method collectionCalculus which can support complex series of " & _
    "operations. " & _
    vbNewline & vbNewline & _
    "This class was created on July 5 2003 by " & vbNewline & vbNewline & _
    "Edward G. Nilges" & vbNewline & _
    "spinoza1111@yahoo.COM" & vbNewline & _
    "http://members.screenz.com/edNilges" & vbNewline & vbNewline & vbNewline & _
    "ISSUES IN COLLECTIONUTILITIES" & vbNewline & vbNewline & _
    "Collection methods are slow for large collections. " & vbNewline & vbNewline & _
    "This class has not been tested completely up to the stress " & _
    "test. All the ""collection calculus"" tests supplied in " & _
    "the collectionTest utility work."
    ' --- Testing
    Private Const STRESS_TEST_COUNT As Integer = 100
    Private Const STRESS_TEST_STEP As String = _
    "rem(""TEST MKRANDOMCOLLECTION AND COLLECTION2REPORT""): " & _ 
    "collection2Report(mkRandomCollection(1,10,""String"",8, _test1)): " & _ 
    "rem(""TEST LIST2COLLECTION AND COLLECTION2SET""): " & _ 
    "collection2Set(list2Collection(1,2,3,3,4,4,5)): Rem: " & _ 
    "rem(""TEST COLLECTION2STRING/COLLECTION2STRINGREADABLE AND STRING2COLLECTION""): " & _ 
    "mkRandomCollection(1,10,""Integer Single String"",8, _test1): " & _ 
    "string2Collection(collection2String(_test1)): " & _ 
    "collection2StringReadable(_test1)): rem: " & _ 
    "rem(""TEST COLLECTIONAPPEND""): " & _ 
    "mkRandomCollection(1, 10, ""Integer Double String Collection"" ,8 , _test2): " & _ 
    "collectionAppend(_test1, _test2): rem: " & _ 
    "rem(""TEST COLLECTIONCLEAR""): " & _ 
    "collectionClear(_test1): rem: " & _ 
    "rem(""TEST COLLECTIONCOPY/COMPARE""): " & _ 
    "collectionCopy(_test1,_test2): " & _ 
    "collectionCompare(_test1,_test2): rem: " & _ 
    "rem(""TEST SET OPERATIONS""): " & _ 
    "mkRandomCollection(1, 10, ""Integer Double String Collection"" ,8 , _test1): " & _ 
    "mkRandomCollection(1, 10, ""Integer Double String Collection"" ,8 , _test2): " & _ 
    "collectionCopy(_test2, _test3): " & _ 
    "collectionComplement(collectionIntersection(collectionUnion(_test1,_test2), _test3), _test1): rem: " & _ 
    "rem(""TEST COLLECTION DEPTH""): " & _ 
    "collectionDepth: Rem: " & _ 
    "rem(""TEST COLLECTION FIND""): " & _ 
    "collectionFind(list2Collection(a,b,c), ""c""): Rem: " & _ 
    "rem(""TEST COLLECTION SAVE""): " & _ 
    "mkRandomCollection(1, 10, ""Integer Double String Collection"" ,8): " & _ 
    "collectionSave(_test1): Rem: " & _ 
    "rem(""TEST COLLECTION TYPES""): " & _ 
    "collectionTypes(_test1): Rem"  
    ' --- mkRandomCollection defaults
    Private Const DEFAULT_RANDOM_MIN As Integer = 0
    Private Const DEFAULT_RANDOM_MAX As Integer = 100
    Private Const DEFAULT_RANDOM_TYPES As String = _
            "Boolean Byte Short Integer Long Single Double String Collection collectionUtilities"
    Private Const DEFAULT_RANDOM_DEPTH As Integer = 8
    
    ' ***** Events *****
    Public Event collection2SetEvent(ByVal strObjectName As String, _
                                     ByVal intItemCount As Integer, _
                                     ByVal intItemIndex As Integer, _
                                     ByVal intLevel As Integer)
    Public Event collection2StringEvent(ByVal strObjectName As String, _
                                        ByVal intItemCount As Integer, _
                                        ByVal intItemIndex As Integer, _
                                        ByVal intLevel As Integer)
    Public Event collectionAppendEvent(ByVal strObjectName As String, _
                                       ByVal intItemCount As Integer, _
                                       ByVal intItemIndex As Integer)
    Public Event collectionCalculusEvent(ByVal strObjectName As String, _
                                         ByVal strOp As String, _
                                         ByVal objResult As Object)
    Public Event collectionClearEvent(ByVal strObjectName As String, _
                                      ByVal intItemCount As Integer, _
                                      ByVal intItemIndex As Integer, _
                                      ByVal intLevel As Integer)
    Public Event collectionCompareEvent(ByVal strObjectName As String, _
                                        ByVal intItemCount As Integer, _
                                        ByVal intItemIndex As Integer, _
                                        ByVal intLevel As Integer)
    Public Event collectionComplementEvent(ByVal strObjectName As String, _
                                           ByVal intItemCount As Integer, _
                                           ByVal intItemIndex As Integer, _
                                           ByVal intLevel As Integer)
    Public Event collectionCopyEvent(ByVal strObjectName As String, _
                                     ByVal intItemCount As Integer, _
                                     ByVal intItemIndex As Integer, _
                                     ByVal intLevel As Integer)
    Public Event collectionDepthEvent(ByVal strObjectName As String, _
                                      ByVal intItemCount As Integer, _
                                      ByVal intItemIndex As Integer, _
                                      ByVal intLevel As Integer)
    Public Event collectionFindEvent(ByVal strObjectName As String, _
                                     ByVal intItemCount As Integer, _
                                     ByVal intItemIndex As Integer, _
                                     ByVal intLevel As Integer)
    Public Event collectionIntersectionEvent(ByVal strObjectName As String, _
                                             ByVal intItemCount As Integer, _
                                             ByVal intItemIndex As Integer, _
                                             ByVal intLevel As Integer)
    Public Event collectionItemSetEvent(ByVal strObjectName As String, _
                                        ByVal strDewey As String, _
                                        ByVal objOldValue As Object, _
                                        ByVal objNewValue As Object)
    Public Event collectionParseEvent(ByVal strObjectName As String, _
                                      ByVal intIndex As Integer, _
                                      ByVal strGoal As String)
    Public Event collectionTypeEvent(ByVal strObjectName As String, _
                                     ByVal intItemCount As Integer, _
                                     ByVal intItemIndex As Integer, _
                                     ByVal intLevel As Integer)
    Public Event collectionUnionEvent(ByVal strObjectName As String, _
                                      ByVal intItemCount As Integer, _
                                      ByVal intItemIndex As Integer, _
                                      ByVal intLevel As Integer)
    Public Event mkRandomCollectionEvent(ByVal strObjectName As String, _
                                         ByVal intItemCount As Integer, _
                                         ByVal intItemIndex As Integer, _
                                         ByVal intLevel As Integer)
    Public Event string2CollectionEvent(ByVal strObjectName As String, _
                                        ByVal intItemCount As Integer, _
                                        ByVal intItemIndex As Integer, _
                                        ByVal intLevel As Integer)
    Public Event testEvent(ByVal strObjectName As String, _
                           ByVal intItemCount As Integer, _
                           ByVal intItemIndex As Integer)
                                        
#End Region ' General Declarations

#Region " Object constructor "      
                                  
    ' ***** Object constructor *********************************************
    
    Public Sub new
        With USRstate
            Interlocked.Increment(_INTsequence)
            .strName = ClassName & _
                       _OBJutilities.alignRight(CStr(_INTsequence), 4, strFill:="0") & " " & _
                       Now
            .booCompareAllowsClone = True                       
            .booUsable = True
        End With        
    End Sub       
    
#End Region ' Object constructor   

#Region " Public procedures (excludes collectionCalculus) "
    
    ' ***** Public procedures **********************************************
    ' *                                                                    *
    ' * Excludes the collectionCalculus code which is in its own region at *
    ' * the end of this class.                                             *
    ' *                                                                    *
    ' * May also contain Private procedures, that perform specific tasks   *
    ' * on behalf of one Public procedure.                                 *
    ' *                                                                    *
    ' **********************************************************************
    
    ' ----------------------------------------------------------------------
    ' Easter Egg
    '
    '
    Public Shared ReadOnly Property About As String
        Get
            Return(CLASS_NAME & vbNewline & vbNewline & ABOUT_INFO)
        End Get        
    End Property    
    
    ' ----------------------------------------------------------------------
    ' Class name
    '
    '
    Public Shared  ReadOnly Property ClassName As String
        Get
            Return(CLASS_NAME)
        End Get        
    End Property    
    
    ' ----------------------------------------------------------------------
    ' Collection to columnar report
    '
    '    
    ' This method produces a report in columns showing the information in
    ' any collection.
    '
    ' This method can display two types of collections:
    '
    '
    '      *  Any collection containing items, none of which is a collection,
    '         is converted to a single column report.
    '
    '         Items having value types (Boolean, Byte, Short, Integer, Long,
    '         Single, Double or String) are converted to their string value
    '         by default (see below for the use of the booDeco parameter to
    '         add type information.)
    '
    '         Items, other than collections, having reference types are
    '         converted to the string value returned by utilities.object2String.
    '
    '      *  Any collection that contains a collection in each entry
    '         is converted to a multiple column report, where each column
    '         contains the subcollection item.  
    '
    ' 
    ' In the report, fields of numeric type are right aligned while fields
    ' of string or Boolean type are left aligned.
    '
    ' By default the columns are displayed without headers.  However, the
    ' optional parameter strColHeaders may be present and it may specify
    ' column headers as a comma-delimited list of words or phrases.
    '
    ' The intGutter optional parameter may specify the size in spaces of 
    ' the "gutter" the area between columns: by default intGutter is 1.
    '
    ' By default, numbers and strings are converted to strings and
    ' reference objects are converted to their representation as returned
    ' by object2String.  Use the optional booDeco:=True parameter to
    ' display each field in the "decorated" format type(value).
    '
    ' Note: output strings built with this method are best viewed in a 
    ' monospaced font such as Courier New.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE   PROGRAMMER     D E S C R I P T I O N
    ' -------- ----------     -----------------------------------------
    ' 01 12 03 Nilges         1.  Version 1.0
    '
    ' 05 06 03 Nilges         1.  Bug: conversion of Nothing
    '
    '
    Public Function collection2Report(ByVal colCollection As Collection, _
                                      Optional ByVal strColHeaders As String = "", _
                                      Optional ByVal intGutter As Integer = 1, _
                                      Optional ByVal booDeco As Boolean = False) _
           As String
        If Not checkUsable_("collection2Report", "Returning a null string") Then
            Return ("")
        End If
        If (colCollection) Is Nothing Then
            errorHandler_("Collection is null", _
                          "collection2Report", _
                          "Returning a null string")
            Return ("")
        End If
        If intGutter < 0 Then
            errorHandler_("Invalid intGutter value is " & intGutter, _
                          "collection2Report", _
                          "Returning a null string")
            Return ("")
        End If
        Dim intIndex1 As Integer
        Dim strColHeaderArray() As String
        Try
            strColHeaderArray = Split(strColHeaders, ",")
        Catch
            errorHandler_("Cannot convert strColHeaders into an array", _
                          "collection2Report", _
                          Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                          "Returning a null string", _
                          True)
            Return ("")
        End Try
        Dim intMaxWidth() As Integer
        Dim strCollection(,) As String
        Try
            ReDim intMaxWidth(UBound(strColHeaderArray) + 1)
            ReDim strCollection(colCollection.Count, UBound(intMaxWidth))
        Catch
            errorHandler_("Cannot create the strCollection array and/or intMaxWidth array", _
                          "collection2Report", _
                          Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                          "Returning a null string", _
                          True)
            Return ("")
        End Try
        For intIndex1 = 1 To UBound(strColHeaderArray)
            intMaxWidth(intIndex1) = Len(strColHeaderArray(intIndex1 - 1))
        Next intIndex1
        Dim booReferenceType As Boolean
        Dim colSubcollection As Collection
        Dim intIndex2 As Integer
        Dim strNext As String
        With colCollection
            ' --- Pass one: determine max column widths and verify we have a usable collection
            For intIndex1 = 1 To .Count
                booReferenceType = _OBJutilities.hasReferenceType(.Item(intIndex1))
                If booReferenceType Then
                    If (TypeOf (.Item(intIndex1)) Is Collection) Then
                        colSubcollection = CType(.Item(intIndex1), Collection)
                    Else
                        booReferenceType = False
                        strNext = _OBJutilities.object2String(.Item(intIndex1))
                    End If
                Else
                    If booDeco Then
                        strNext = _OBJutilities.object2String(.Item(intIndex1), True)
                    Else
                        strNext = CStr(.Item(intIndex1))
                    End If
                End If
                If Not booReferenceType Then
                    Try
                        colSubcollection = New Collection
                        colSubcollection.Add(strNext)
                    Catch
                        errorHandler_("Cannot create temporary sub-collection", _
                                      "collection2Report", _
                                      Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                                      "Returning a null string", _
                                      True)
                        Return ("")
                    End Try
                End If
                With colSubcollection
                    If .Count > UBound(intMaxWidth) Then
                        Try
                            ReDim Preserve strCollection(UBound(strCollection), .Count)
                            ReDim Preserve intMaxWidth(.Count)
                        Catch
                            errorHandler_("Cannot expand intMaxWidth collection:" & _
                                          Err.Number & " " & Err.Description, _
                                          "collection2Report", _
                                          "Returning a null string", _
                                          True)
                            Return ("")
                        End Try
                    End If
                    For intIndex2 = 1 To UBound(strCollection, 2)
                        Try
                            If booDeco Then
                                strCollection(intIndex1, intIndex2) = _
                                    _OBJutilities.object2String(.Item(intIndex2), booDeco)
                            Else
                                strCollection(intIndex1, intIndex2) = CStr(.Item(intIndex2))
                            End If
                        Catch
                            errorHandler_("Cannot convert collection entry " & _
                                          _OBJutilities.object2String(.Item(intIndex2)) & " " & _
                                          "at row " & intIndex1 & " " & _
                                          "and column " & intIndex2 & " " & _
                                          "to a string", _
                                          "collection2Report", _
                                          Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                                          "Returning a null string")
                            Return ("")
                        End Try
                        intMaxWidth(intIndex2) = Math.Max(intMaxWidth(intIndex2), _
                                                          Len(strCollection(intIndex1, intIndex2)))
                    Next intIndex2
                End With
                If Not booReferenceType Then colSubcollection = Nothing
            Next intIndex1
            ' --- Pass two: create the report
            Dim objStringBuilder As System.Text.StringBuilder
            Try
                objStringBuilder = New System.Text.StringBuilder
            Catch
                errorHandler_("Cannot create StringBuilder", _
                              "collection2Report", _
                              Err.Number & Err.Description & vbNewLine & vbNewLine & _
                              "Returning a null string", _
                              True)
                Return ("")
            End Try
            objStringBuilder.Append(vbNewLine)
            Dim strGutter As String = _OBJutilities.copies(" ", intGutter)
            Dim strGutter2 As String
            If Trim(strColHeaders) <> "" Then
                For intIndex1 = 0 To UBound(strColHeaderArray)
                    _OBJutilities.append(objStringBuilder, _
                                         strGutter2, _
                                         _OBJutilities.alignCenter(strColHeaderArray(intIndex1), _
                                         intMaxWidth(intIndex1 + 1)))
                    strGutter2 = strGutter
                Next intIndex1
                objStringBuilder.Append(vbNewLine)
                strGutter2 = ""
                For intIndex1 = 0 To UBound(strColHeaderArray)
                    _OBJutilities.append(objStringBuilder, _
                                         strGutter2, _
                                         _OBJutilities.copies("-", _
                                                              intMaxWidth(intIndex1 + 1)))
                    strGutter2 = strGutter
                Next intIndex1
            End If
            Dim objStringBuilder2 As System.Text.StringBuilder
            Try
                objStringBuilder2 = New System.Text.StringBuilder
            Catch
                errorHandler_("Cannot create string builder for row display", _
                              "collection2Report", _
                              Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                              "Returning a null string", _
                              True)
                Return ("")
            End Try
            For intIndex1 = 1 To UBound(strCollection, 1)
                For intIndex2 = 1 To UBound(strCollection, 2)
                    If IsNumeric(strCollection(intIndex1, intIndex2)) Then
                        strCollection(intIndex1, intIndex2) = _
                            _OBJutilities.alignRight(strCollection(intIndex1, intIndex2), intMaxWidth(intIndex2))
                    Else
                        strCollection(intIndex1, intIndex2) = _
                            _OBJutilities.alignLeft(strCollection(intIndex1, intIndex2), intMaxWidth(intIndex2))
                    End If
                    _OBJutilities.append(objStringBuilder2, _
                                            strGutter, _
                                            strCollection(intIndex1, intIndex2))
                Next intIndex2
                _OBJutilities.append(objStringBuilder, vbNewLine, objStringBuilder2.ToString)
                objStringBuilder2.Length = 0
            Next intIndex1
            Return (objStringBuilder.ToString & vbNewLine)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Convert a collection to a set
    '
    '
    Public Function collection2Set(ByRef colCollection As Collection) As Boolean
        If Not checkUsable_("collection2Set", "Collection not changed") Then
            Return (False)
        End If
        Return (collection2Set_(colCollection, colCollection, ""))
    End Function

    ' ---------------------------------------------------------------------
    ' Recursion on behalf of collection2Set
    '
    '
    Private Function collection2Set_(ByRef colCollection As Collection, _
                                     ByVal colMain As Collection, _
                                     ByVal strDewey As String) As Boolean
        If Not checkUsable_("collection2Set", "Collection not changed") Then
            Return (False)
        End If
        If (colCollection Is Nothing) Then Return (True)
        Dim intLevel As Integer = _OBJutilities.items(strDewey, ".", False)
        Dim strInnerDeweyPrefix As String = strDewey & CStr(IIf(strDewey = "", "", "."))
        Dim strInnerDewey As String
        Dim strKey As String
        With colCollection
            Dim intIndex1 As Integer
            For intIndex1 = .Count To 1 Step -1
                RaiseEvent collection2SetEvent(Me.Name, .Count, intIndex1, intLevel)
                strInnerDewey = strInnerDeweyPrefix & intIndex1
                strKey = CStr(Me.collectionFind(colMain, .Item(intIndex1)))
                If strKey <> "0" _
                   AndAlso _
                   strKey <> strInnerDewey Then
                    Try
                        .Remove(intIndex1)
                    Catch
                        errorHandler_("Could not remove duplicate element: " & _
                                      Err.Number & " " & Err.Description, _
                                      "collection2Set_", _
                                      "Returning False, but collection may have been modified in part")
                        Exit For
                    End Try
                ElseIf (TypeOf .Item(intIndex1) Is Collection) Then
                    If Not collection2Set_(CType(.Item(intIndex1), Collection), _
                                           colCollection, _
                                           strInnerDewey) Then Exit For
                End If
            Next intIndex1
            Return (intIndex1 < 1)
        End With
    End Function

    ' ---------------------------------------------------------------------
    ' Convert a collection to a string
    '
    ' 
    ' This function accepts a Collection and returns the comma-delimited
    ' string of its elements. This value can be returned optimized for
    ' readable display or it can be returned optimized for accurate
    ' serialization, including the ability to rebuild the identical collection.
    ' This feature makes this method useful in XML applications!
    '
    ' This documentation header discusses the following topics.
    '
    '
    '      *  Returning readable strings
    '      *  Returning serialized collections
    '      *  Certifying serialized collections: palindrome theory
    '      *  Inspection
    '       
    '
    ' RETURNING READABLE STRINGS
    '
    ' By default, the simple syntax collection2String(collection) returns
    ' a readable collection that will not convert back to the same
    ' collection. The following strings are returned.
    '
    '
    '      *  If the input collection is Nothing, the string noCollection is
    '         returned
    '
    '      *  If the input collection is empty, the string emptyCollection is
    '         returned
    '
    '      *  If the input collection is allocated and also non-empty:
    ' 
    '         + Entries that are Nothing are converted to the string Nothing
    ' 
    '         + Numeric members are included without change
    '
    '         + Nonblank, non-null strings that contain only letters, numbers, 
    '           and graphic special characters (characters available on common 
    '           USA keyboards and common fonts), other than parentheses or commas, 
    '           are included without change
    '
    '         + Blank and null strings are returned in quotes
    '
    '         + Other strings are encoded as XML strings: nongraphic characters,
    '           commas and parentheses are replaced by &#nnnnn where nnnnn is 
    '           the XML character value
    '
    '         + Collections, embedded in the main collection, are
    '           included as the parenthesized output of this function,
    '           called recursively.  
    '
    '           The recursion depth is limited to the value in the optional
    '           parameter intRecursionDepth.  When recursion depth is exceeded,
    '           contained collections are returned as the word Collection only.  This
    '           avoids the possibility of a loop when a collection contains
    '           a reference to a container.  If intRecursionDepth is -1 then
    '           there is no limit to recursion.  The default value of this
    '           parameter is 8 levels.
    '
    '         + All other object members are converted using object2String with
    '           decoration which will usually return Object, the object type name 
    '           or Nothing: see the object2String method for details. 
    '
    '
    ' Suppose, then, that a collection contains these members:
    '
    '
    '      1.  Item 1 is the string "Member 1 of 3"
    '      2.  Item 2 is the integer 2
    '      3.  Item 3 is a collection and it contains these members:
    '          3.1 Item 1 is Nothing
    '          3.2 Item 2 is a collection containing one member set to the
    '              Boolean value True
    '      4.  Item 4 is a stringBuilder object
    '
    '
    ' By default, collection2String(collection) will return this readable string:
    '
    '
    '      Member 1 of 3,2,(Nothing,(True)),System.Text.StringBuilder
    '
    '
    ' In addition, the booDeco optional parameter may be used with booReadable:=True;
    ' this will cause each element to be returned in the form type(value), where
    ' type is the type of the element and value is its value. The value will be a number,
    ' a quoted string, or the toString result of a reference object when available;
    ' in other cases it will be Object.
    '
    '
    ' RETURNING SERIALIZED COLLECTIONS USING BOOREADABLE:=FALSE
    '
    ' The optional parameter booReadable:=False tries to return a less-readable string
    ' that accurately preserves the collection in such a way that string2Collection
    ' can return the identical collection (a clone) from the "serialized" collection
    ' when the original collection contains value objects and collections exclusively. 
    ' Here, each entry has the format Nothing, type(value), just (value) or Object:
    '
    '
    '      *  For Nothing and empty collections, noCollection and emptyCollection
    '         are returned as for readable strings, as above
    '
    '      *  For allocated collections with members:
    '
    '         + For a collection entry of Nothing, the string Nothing is returned
    '
    '         + For a numeric type, type is the system type name and value is
    '           the value: for example, System.Byte(255)
    '
    '         + For a string, type is System.String and value is the quoted
    '           string value: for example, System.String("a"). In the quoted
    '           string value, nongraphic characters, quotes, parentheses and the
    '           comma are replaced by their XML equivalents &#nnnnn.
    '
    '         + For a contained collection that has been successfully serialized, the
    '           (value) is present such that the parenthesized value is the
    '           contained collection, serialized according to the rules of this
    '           method (including booReadable:=False).
    '
    '         + For a collection that has NOT been successfully serialized and for
    '           all reference objects other than collections, type(value), type, or
    '           Object is returned, according to the rules of object2String. The 
    '           type will be the system name of the type. For a collection this 
    '           will be Microsoft.VisualBasic.Collection.
    '
    '
    ' When booReadable:=False is supplied collection2String will return this string
    ' for the above example collection.
    '
    '
    '      System.String("Member 1 of 4"),
    '      System.Int32(2),
    '      (Nothing,(System.Boolean(True))),
    '      System.Text.StringBuilder
    '
    '
    ' Note: when booReadable is False, booDeco will have no effect.
    '
    '
    ' CERTIFYING SERIALIZED COLLECTIONS: PALINDROME THEORY
    '
    ' Note that the output of this method is either a "palindrome" or "nonpalindromic."
    ' A palindrome is a string or other object that is reversible such as the word
    ' "aha". For our purposes, the output of this method is a palindrome when it is
    ' known that string2Object will create a clone collection from the output string.
    '
    ' No readable string is a palindrome since string2Collection is unable to convert
    ' their syntax; to be a palindrome, the collection2String result has to contain
    ' the above object type and value information. Even if it does, if the collection2String
    ' result contains collections that were not expanded to strings, or reference objects
    ' other than collections, then it is not a palindrome.
    '
    ' To both create a palindrome use booReadable:=False. To obtain certification that a 
    ' palindrome was returned, use the overload syntax collection2String(colCollection, 
    ' booPalindrome, booReadable:=False) where booPalindrome is a Boolean passed by 
    ' reference. booPalindrome will be set to True if the string returned is a palindrome, 
    ' False otherwise.
    '
    ' By default, the individual collection items are separated by commas. The optional
    ' strSeparator parameter may define the separator for all levels, or the optional
    ' strSeparator1 parameter may define only the top-level separator that's used to
    ' separate only items in the collection passed at its top level. If both strSeparator
    ' and strSeparator1 are specified, then strSeparator1 is used at the top level, while
    ' strSeparator is used for all other separators.
    '
    ' Note that if either strSeparator1 or strSeparator are not set to their default
    ' values, the result is not palindromic.
    '
    '
    ' INSPECTION
    '
    ' By default, when booReadable is False and the string form of the collection is a 
    ' palindrome, this string is inspected for correctness, by converting it back to a 
    ' new collection, and ensuring that this collection is a clone of the original 
    ' collection, using our collectionCompare method. An error is raised if the 
    ' collections don't match.
    '
    ' To suppress inspection (saving time) use the optional parameter booInspect:=False.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE   PROGRAMMER     D E S C R I P T I O N
    ' -------- ----------     -----------------------------------------
    ' 11 20 02 Nilges         1.  Version 1.0
    ' 02 17 03 Nilges         1.  Added strDelim parameter
    '                         2.  Decorate the string
    ' 06 02 03 Nilges         1.  Use XML
    '                         2.  Added booReadable and booPalindrome.
    '                         3.  Removed strDelim
    ' 07 02 03 Nilges         1.  Added booInspect
    ' 07 04 03 Nilges         1.  Unreadable strings should be in the
    '                             format of object2String
    ' 07 16 03 Nilges         1.  Never inspect when readability is in
    '                             effect
    ' 09 29 03 Nilges         1.  Added booDeco parameter
    ' 02 07 04 Nilges         1.  Added strSeparator parameters
    '
    '
    Public Overloads Function collection2String(ByVal colCollection As Collection, _
                                                Optional ByVal intRecursionDepth As Integer = 8, _
                                                Optional ByVal booReadable As Boolean = False, _
                                                Optional ByVal booInspect As Boolean = True, _
                                                Optional ByVal booDeco As Boolean = True, _
                                                Optional ByVal strSeparator1 As String = ",", _
                                                Optional ByVal strSeparator As String = ",") _
           As String
        Dim booPalindrome As Boolean
        Return (collection2String(colCollection, _
                                 booPalindrome, _
                                 intRecursionDepth:=intRecursionDepth, _
                                 booReadable:=booReadable, _
                                 booInspect:=booInspect, _
                                 booDeco:=booDeco, _
                                 strSeparator:=strSeparator, _
                                 strSeparator1:=strSeparator1))
    End Function
    Public Overloads Function collection2String(ByVal colCollection As Collection, _
                                                ByRef booPalindrome As Boolean, _
                                                Optional ByVal intRecursionDepth As Integer = 8, _
                                                Optional ByVal booReadable As Boolean = False, _
                                                Optional ByVal booInspect As Boolean = True, _
                                                Optional ByVal booDeco As Boolean = True, _
                                                Optional ByVal strSeparator1 As String = ",", _
                                                Optional ByVal strSeparator As String = ",") _
           As String
        If Not checkUsable_("collection2Report", "Returning a null string") Then
            Return ("")
        End If
        booPalindrome = False
        If intRecursionDepth < -1 Then
            errorHandler_("Invalid recursion depth limit of " & intRecursionDepth, _
                          "collection2String", _
                          "This parameter must be zero, positive or -1 for no limit")
            Return ("")
        End If
        If (colCollection Is Nothing) Then
            booPalindrome = Not booReadable
            Return ("noCollection")
        End If
        If (colCollection.Count = 0) Then
            booPalindrome = Not booReadable
            Return ("emptyCollection")
        End If
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler_("Cannot create stringBuilder", _
                          "collection2String", _
                          Err.Number & " " & Err.Description, _
                          True)
            Return ("")
        End Try
        Dim intIndex1 As Integer
        Dim intRecursion As Integer
        Dim objNext As Object
        Dim strNext As String
        Dim strNextWork As String
        Dim booPalindromeInner As Boolean
        booPalindrome = Not booReadable
        For intIndex1 = 1 To colCollection.Count
            objNext = colCollection.Item(intIndex1)
            If (TypeOf objNext Is Collection) Then
                ' --- Inner collection
                If intRecursion <> -1 AndAlso intRecursion >= intRecursionDepth Then
                    ' Can't serialize
                    If booReadable Then
                        strNext = "Collection"
                    Else
                        strNext = _OBJutilities.object2String(objNext)
                    End If
                    booPalindrome = False
                Else
                    ' Try to serialize collection
                    strNext = collection2String(CType(objNext, Collection), _
                                                intRecursionDepth:=intRecursion + 1, _
                                                booPalindrome:=booPalindromeInner, _
                                                booReadable:=booReadable, _
                                                booInspect:=False, _
                                                strSeparator1:=strSeparator, _
                                                strSeparator:=strSeparator)
                    booPalindrome = (booPalindrome AndAlso booPalindromeInner)
                    If Not booReadable Then
                        strNextWork = strNext
                        strNext = _OBJutilities.string2Display(strNext, _
                                                               "XML", _
                                                               strGraphicExclude:=",()")
                        booPalindrome = (booPalindrome AndAlso (strNextWork = strNext))
                    End If
                    strNext = "(" & strNext & ")"
                End If
            Else
                ' --- Well, it's not a collection
                If booReadable Then
                    If booDeco Then
                        strNext = _OBJutilities.object2String(objNext)
                    Else
                        If (objNext Is Nothing) Then
                            strNext = "Nothing"
                        Else
                            Try
                                strNext = CStr(objNext)
                                If Trim(strNext) = "" Then
                                    strNext = _OBJutilities.enquote(strNext)
                                Else
                                    strNext = _OBJutilities.string2Display(CStr(objNext), _
                                                                            "XML", _
                                                                            strGraphicExclude:=",()")
                                End If
                            Catch
                                strNext = _OBJutilities.object2String(objNext, True)
                            End Try
                        End If
                    End If
                Else
                    ' Need to return type(value) to preserve palindrome info
                    If (TypeOf objNext Is System.String) Then
                        ' We need to get rid of characters that may confuse string2Collection
                        strNext = objNext.GetType.ToString & _
                                  "(" & _
                                  _OBJutilities.enquote(_OBJutilities.string2Display(CStr(objNext), _
                                                                                     strGraphicExclude:=",()""")) & _
                                  ")"
                    Else
                        strNext = _OBJutilities.object2String(objNext, True)
                    End If
                End If
            End If
            _OBJutilities.append(objStringBuilder, strSeparator1, strNext)
            RaiseEvent collection2StringEvent(Me.Name, colCollection.Count, intIndex1, intRecursionDepth)
        Next intIndex1
        Dim strOutstring As String = objStringBuilder.ToString
        If Not booReadable AndAlso booPalindrome AndAlso booInspect Then
            Dim colInspect As Collection = string2Collection(strOutstring, booInspect:=False)
            Dim strExplanation As String = "Conversion of output string back to collection has failed"
            If (colInspect Is Nothing) _
               OrElse _
               Not collectionCompare(colCollection, colInspect, strExplanation) Then
                errorHandler_("Inspection failure: " & strExplanation, _
                              "collection2String", _
                              "Returning the converted collection string. It probably is in error.", _
                              True)
            End If
        End If
        Return strOutstring
    End Function

    ' -----------------------------------------------------------------
    ' Convert collection to XML
    '
    '
    ' --- Tag name will be anonymousCollection: subcollections will be
    ' --- converted
    Public Overloads Function collection2XML(ByVal colCollection As Collection) _
           As String
        Return collection2XML_(colCollection, _
                               "anonymousCollection", _
                               True, _
                               True, _
                               0)
    End Function
    ' --- Tag name supplied: subcollections will be converted
    Public Overloads Function collection2XML(ByVal colCollection As Collection, _
                                             ByVal strTagName As String) _
           As String
        Return collection2XML_(colCollection, _
                               strTagName, _
                               True, _
                               True, _
                               0)
    End Function
    ' --- Tag name supplied: subcollection conversion specified
    Public Overloads Function collection2XML(ByVal colCollection As Collection, _
                                             ByVal strTagName As String, _
                                             ByVal booRecursion As Boolean) _
           As String
        Return collection2XML_(colCollection, _
                               strTagName, _
                               booRecursion, _
                               True, _
                               0)
    End Function
    ' --- Tag name supplied: subcollection conversion specified
    Public Overloads Function collection2XML(ByVal colCollection As Collection, _
                                             ByVal strTagName As String, _
                                             ByVal booRecursion As Boolean, _
                                             ByVal booFormat As Boolean) _
           As String
        Return collection2XML_(colCollection, _
                               strTagName, _
                               booRecursion, _
                               booFormat, _
                               0)
    End Function
    ' --- Common logic and recursion step
    Private Overloads Function collection2XML_(ByVal colCollection As Collection, _
                                               ByVal strTagName As String, _
                                               ByVal booRecursion As Boolean, _
                                               ByVal booFormat As Boolean, _
                                               ByVal intLevel As Integer) _
           As String
        Dim strOutstring As String = _OBJutilities.mkXMLTag(strTagName)
        If (colCollection Is Nothing) Then
            Return strOutstring & "Nothing" & _OBJutilities.mkXMLTag(strTagName, True)
        End If
        With colCollection
            Dim intIndex1 As Integer
            Dim intItemSeq As Integer = 1
            Dim intSubcollectionSeq As Integer = 1
            Dim strIndent As String = _OBJutilities.copies(" ", (intLevel + 1) * 4)
            Dim strItemTagName As String
            Dim strNewline As String = CStr(IIf(booFormat, vbNewLine, ""))
            Dim strNext As String
            Dim intLength As Integer
            If .Count <> 0 Then intLength = CInt(Math.Log10(.Count) + 1)
            For intIndex1 = 1 To .Count
                strItemTagName = "Item" & _OBJutilities.alignRight(CStr(intItemSeq), intLength, "0")
                If booRecursion AndAlso (TypeOf .Item(intIndex1) Is Collection) Then
                    strNext = strNewline & _
                                collection2XML_(CType(.Item(intIndex1), Collection), _
                                                "subcollection" & intSubcollectionSeq, _
                                                True, _
                                                booFormat, _
                                                intLevel + 1) & _
                                strNewline & strIndent
                    intSubcollectionSeq += 1
                Else
                    strNext = _OBJutilities.object2String(.Item(intIndex1), True)
                End If
                intItemSeq += 1
                If booFormat Then
                    strOutstring &= strNewline & strIndent
                End If
                strOutstring &= _OBJutilities.mkXMLElement(strItemTagName, strNext)
            Next intIndex1
            strOutstring &= strNewline & _OBJutilities.mkXMLTag(strTagName, True)
            If booFormat AndAlso intLevel > 0 Then
                strIndent = _OBJutilities.copies(" ", intLevel * 4)
                strOutstring = strIndent & _
                               Replace(strOutstring, strNewline, strNewline & strIndent)
            End If
            Return strOutstring
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Collection append
    '
    '
    Public Function collectionAppend(ByRef colCollection1 As Collection, _
                                     ByVal colCollection2 As Collection) _
                    As Boolean
        If Not checkUsable_("collectionAppend", _
                            "Collection 1 is not changed") Then
            Return (False)
        End If
        Dim intIndex1 As Integer
        With colCollection2
            For intIndex1 = 1 To .Count
                RaiseEvent collectionAppendEvent(Me.Name, .Count, intIndex1)
                Try
                    colCollection1.Add(.Item(intIndex1))
                Catch
                    errorHandler_("Cannot add item " & intIndex1 & " of collection 2 " & _
                                  "to collection 1: " & _
                                  Err.Number & " " & Err.Description, _
                                  "collectionAppend", _
                                  "Returning False: collection 1 may have been " & _
                                  "partially changed")
                End Try
            Next intIndex1
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Clear collection (including contained objects and collections)
    '
    '
    ' This method removes all items from colCollection. For each item that
    ' is a collection, this method recursively clears that item. It sets 
    ' each reference item to Nothing.
    '
    ' By default, the main colCollection will itself be set to Nothing after
    ' the clear; however, this can be suppressed: when booSetToNothing is
    ' present and False, the main colCollection will exist but contain no
    ' members, on return from this method. 
    '
    ' We need to avoid loops when for some silly reason, a collection
    ' references itself. Therefore, we pass, to each recursion level, a
    ' collection of the "pending" collection handles.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 06 12 03   Nilges         Added booSetToNothing
    '
    '
    Public Function collectionClear(ByRef colCollection As Collection, _
                                    Optional ByVal booSetToNothing As Boolean = True) _
           As Boolean
        If Not checkUsable_("collectionClear", "Not clearing string and returning False") Then
            Return (False)
        End If
        Dim colPending As Collection
        Try
            colPending = New Collection
        Catch
            errorHandler_("Can't create the pending collection: " & _
                          Err.Number & " " & Err.Description, _
                          "collectionClear", _
                          "Cannot clear collection", _
                          True)
            Return (False)
        End Try
        If Not collectionClear_(colCollection, colPending, booSetToNothing, 0) Then Return (False)
        Return (collectionClear_(colPending, Nothing, True, 0))
    End Function
    ' --- Recursion
    Private Function collectionClear_(ByRef colCollection As Collection, _
                                      ByVal colPending As Collection, _
                                      ByVal booSetToNothing As Boolean, _
                                      ByVal intLevel As Integer) As Boolean
        Dim colInnerCollection As Collection
        Dim intIndex1 As Integer
        Dim objItemHandle As Object
        If Not (colPending Is Nothing) Then
            With colPending
                For intIndex1 = 1 To .Count
                    If (.Item(intIndex1) Is colCollection) Then
                        errorHandler_("Collection references itself", _
                                      "collectionClear", _
                                      "Cannot clear collection")
                        Return (False)
                    End If
                Next intIndex1
            End With
        End If
        With colCollection
            For intIndex1 = .Count To 1 Step -1
                objItemHandle = .Item(intIndex1)
                .Remove(intIndex1)
                RaiseEvent collectionClearEvent(Me.Name, .Count, intIndex1, intLevel)
                If (TypeOf objItemHandle Is Collection) Then
                    colInnerCollection = CType(objItemHandle, Collection)
                    Try
                        colPending.Add(colCollection)
                    Catch
                        errorHandler_("Can't push to the pending collection: " & _
                                      Err.Number & " " & Err.Description, _
                                      "collectionClear", _
                                      "Cannot clear collection", _
                                      True)
                        Return (False)
                    End Try
                    If Not collectionClear_(colInnerCollection, _
                                            colPending, _
                                            True, _
                                            intLevel + 1) Then Return (False)
                    Try
                        colPending.Remove(colPending.Count)
                    Catch
                        errorHandler_("Can't pop from the pending collection: " & _
                                      Err.Number & " " & Err.Description, _
                                      "collectionClear", _
                                      "Cannot clear collection", _
                                      True)
                        Return (False)
                    End Try
                Else
                    Try
                        objItemHandle = Nothing
                    Catch : End Try
                End If
            Next intIndex1
            If booSetToNothing Then colCollection = Nothing
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Compare many collections
    '
    '
    ' This method compares two well-behaved collections that are "well-
    ' behaved" in that they contain "canonical" types (Boolean, Byte,
    ' Short, Integer, Long, Single, Double or String), or sub-collections,
    ' restricted to canonical types and sub-sub-collections, in each
    ' entry.
    '
    ' This method returns True (collections are identical) or False.
    '
    ' To obtain an explanation of the first difference found (which will
    ' cause this method to return False) use the overload
    ' collectionCompare(c1, c2, explanation), where explanation is a string
    ' passed by reference.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 06 26 03   Nilges         1.  Bug: first overload was a loop
    '                           2.  Need to use explanation and not
    '                               error handling when collection
    '                               is compared to noncollection
    '
    '
    ' --- Basic call
    Public Overloads Function collectionCompare(ByVal colCollection1 As Collection, _
                                                ByVal colCollection2 As Collection) As Boolean
        If Not checkUsable_("collectionCompare", "Not comparing: returning False") Then
            Return (False)
        End If
        Dim strExplanation As String
        Return (collectionCompare(colCollection1, colCollection2, strExplanation))
    End Function
    ' --- Explain first difference
    Public Overloads Function collectionCompare(ByVal colCollection1 As Collection, _
                                                ByVal colCollection2 As Collection, _
                                                ByRef strExplanation As String) As Boolean
        If Not checkUsable_("collectionCompare", "Not comparing: returning False") Then
            Return (False)
        End If
        Return (collectionCompare_(colCollection1, colCollection2, strExplanation, 0))
    End Function
    ' --- Recursion
    Private Function collectionCompare_(ByVal colCollection1 As Collection, _
                                        ByVal colCollection2 As Collection, _
                                        ByRef strExplanation As String, _
                                        ByVal intLevel As Integer) As Boolean
        With colCollection1
            If .Count <> colCollection2.Count Then
                strExplanation = "The size of collection 1 is " & .Count & " while " & _
                                 "the size of collection 2 is " & colCollection2.Count
                Return (False)
            End If
            Dim colInnerCollection As Collection
            Dim intIndex1 As Integer
            Dim objItemHandle As Object
            For intIndex1 = .Count To 1 Step -1
                RaiseEvent collectionCompareEvent(Me.Name, .Count, intIndex1, intLevel)
                objItemHandle = .Item(intIndex1)
                If (TypeOf objItemHandle Is Collection) Then
                    colInnerCollection = CType(objItemHandle, Collection)
                    objItemHandle = colCollection2.Item(intIndex1)
                    If (TypeOf objItemHandle Is Collection) Then
                        If Not collectionCompare_(colInnerCollection, _
                                                  CType(objItemHandle, Collection), _
                                                  strExplanation, _
                                                  intLevel + 1) Then
                            Return (False)
                        End If
                    Else
                        strExplanation = "Collection is compared to noncollection"
                        Return (False)
                    End If
                Else
                    If Not matchObjects_(.Item(intIndex1), objItemHandle) Then
                        strExplanation = "Object " & _
                                         _OBJutilities.object2String(.Item(intIndex1)) & " " & _
                                         "does not match object " & _
                                         _OBJutilities.object2String(objItemHandle)
                        Return (False)
                    End If
                End If
            Next intIndex1
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Form complement set
    '
    '
    Public Function collectionComplement(ByRef colCollection1 As Collection, _
                                         ByRef colCollection2 As Collection) _
                    As Collection
        If Not checkUsable_("collectionComplement", _
                            "No complement formed") Then Return (Nothing)
        Return collectionComplement_(colCollection1, colCollection2, 0)
    End Function

    ' ----------------------------------------------------------------------
    ' Recursion on behalf of collectionComplement
    '
    '
    Private Function collectionComplement_(ByVal colCollection1 As Collection, _
                                           ByVal colCollection2 As Collection, _
                                           ByVal intLevel As Integer) _
                     As Collection
        Dim colComplement As Collection = Me.list2Collection()
        If (colComplement Is Nothing) Then Return (Nothing)
        Dim intIndex1 As Integer
        With colCollection2
            For intIndex1 = 1 To .Count
                RaiseEvent collectionComplementEvent(Me.Name, .Count, intIndex1, intLevel)
                If (TypeOf .Item(intIndex1) Is Collection) Then
                    Dim colInnerComplement As Collection = _
                    collectionComplement_(CType(.Item(intIndex1), Collection), _
                                          colCollection2, _
                                          intLevel + 1)
                    If Not Me.collectionAppend(colComplement, _
                                               colInnerComplement) Then
                        Me.collectionClear(colComplement)
                        Return (Nothing)
                    End If
                Else
                    If CStr(Me.collectionFind(colCollection1, .Item(intIndex1))) = "0" Then
                        Try
                            colComplement.Add(.Item(intIndex1))
                        Catch
                            errorHandler_("Unable to extend complement: " & _
                                        Err.Number & " " & Err.Number, _
                                        "collectionComplement_", _
                                        "Returning partial result")
                        End Try
                    End If
                End If
            Next intIndex1
            Return (colComplement)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Copy many collections
    '
    '
    ' This method copies two well-behaved collections that are "well-
    ' behaved" in that they contain "canonical" types (Boolean, Byte,
    ' Short, Integer, Long, Single, Double or String), or sub-collections,
    ' restricted to canonical types and sub-sub-collections, in each
    ' entry.
    '
    ' This method copies colCollection if colCollection is well-behaved,
    ' and it returns a new collection. On error, this method returns Nothing.
    '
    ' By default this method checks its work. By default it executes our
    ' collectionCompare method which compares the copy to the original 
    ' collection for equality only. This method Throws an error and returns
    ' Nothing if this check fails.
    '
    ' To suppress this check use the overload collectionCopy(collection, False)
    '
    '
    ' --- Copy with verification
    Public Overloads Function collectionCopy(ByVal colCollection As Collection) As Collection
        Return (collectionCopy(colCollection, True))
    End Function
    ' --- Copy with no verification
    Public Overloads Function collectionCopy(ByVal colCollection As Collection, _
                                             ByVal booVerification As Boolean) As Collection
        If Not checkUsable_("collectionCompare", "Not copying: returning Nothing") Then
            Return (Nothing)
        End If
        Dim colCopy As Collection = collectionCopy_(colCollection, 0)
        If (colCopy Is Nothing) Then Return (Nothing)
        If booVerification Then
            Dim strExplanation As String
            If Not collectionCompare(colCollection, colCopy, strExplanation) Then
                errorHandler_("Internal programming error: copied collection " & _
                              "doesn't match the original collection" & _
                              vbNewLine & vbNewLine & _
                              strExplanation & _
                              vbNewLine & vbNewLine & _
                              "The original collection is " & _
                              _OBJutilities.enquote(_OBJutilities.ellipsis(collection2String(colCollection), 512)) & _
                              vbNewLine & _
                              "The copy is " & _
                              _OBJutilities.enquote(_OBJutilities.ellipsis(collection2String(colCopy), 512)), _
                              "collectionCopy", _
                              "Destroying the copy and returning Nothing", _
                              True)
                Try
                    collectionClear(colCopy)
                    Return (Nothing)
                Catch : End Try
            End If
        End If
        Return (colCopy)
    End Function
    ' --- Recursion
    Private Function collectionCopy_(ByVal colCollection As Collection, _
                                     ByVal intLevel As Integer) As Collection
        With colCollection
            Dim colCopy As Collection
            Try
                colCopy = New Collection
            Catch ex As Exception
                errorHandler_("Cannot create copy collection: " & _
                              Err.Number & " " & Err.Description, _
                              "collectionCopy_", _
                              "Returning Nothing", _
                              True)
                Return (Nothing)
            End Try
            Dim intIndex1 As Integer
            Dim objItemHandle As Object
            For intIndex1 = 1 To .Count
                RaiseEvent collectionCopyEvent(Me.Name, .Count, intIndex1, intLevel)
                objItemHandle = .Item(intIndex1)
                If (TypeOf objItemHandle Is Collection) Then
                    Dim colInnerCollection As Collection = _
                        collectionCopy_(CType(objItemHandle, Collection), intLevel + 1)
                    If (colInnerCollection Is Nothing) Then
                        errorHandler_("Cannot copy inner collection", _
                                      "collectionCopy_", _
                                      "Destroying the copy and returning Nothing", _
                                      True)
                        Try
                            collectionClear(colCopy)
                            Return (Nothing)
                        Catch : End Try
                    End If
                    objItemHandle = colInnerCollection
                Else
                    Try
                        Dim strScalar As String = CStr(objItemHandle)
                    Catch
                        Try
                            objItemHandle = collectionUtilitiesX.cloneAttempt(objItemHandle)
                        Catch
                            errorHandler_("Collection item types are not supported", _
                                        "collectionCompare", _
                                        "Destroying the copy and returning Nothing")
                            Try
                                collectionClear(colCopy)
                                Return (Nothing)
                            Catch : End Try
                        End Try
                    End Try
                End If
                Try
                    colCopy.Add(objItemHandle)
                Catch
                    errorHandler_("Cannot copy inner collection: " & _
                                  Err.Number & " " & Err.Description, _
                                  "collectionCompare", _
                                  "Destroying the copy and returning Nothing", _
                                  True)
                    Try
                        collectionClear(colCopy)
                        Return (Nothing)
                    Catch : End Try
                End Try
            Next intIndex1
            Return (colCopy)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return collection depth
    '
    '
    ' The "depth" of a collection is the maximum level at which it contains
    ' items which are sub-collections. If the collection is Nothing, its
    ' depth is returned as 0 without error. If the collection contains no
    ' subcollections its depth is 1. Otherwise, the depth is recursively
    ' 1 plus the depth of the deepest subcollection. For example:
    '
    '
    '      *  The depth of the collection [1,2,3] is 1 because it contains
    '         no subcollections    
    '
    '      *  The depth of the collection [1,2,("a"),3] is 2 because it 
    '         contains the subcollection ("a")    
    '
    '      *  The depth of the collection [1,2,("a",(("b"))),3] is 4 because 
    '         it contains the subcollection ("a",(("b"))), and this
    '         subcollection has depth 2.    
    '
    '
    ' --- Public interface
    Public Overloads Function collectionDepth(ByVal colCollection As Collection) As Integer
        Return (collectionDepth(colCollection, 0))
    End Function
    ' --- Recursion
    Private Overloads Function collectionDepth(ByVal colCollection As Collection, _
                                               ByVal intLevel As Integer) As Integer
        If Not checkUsable_("collectionDepth", "Returning 0") Then
            Return (0)
        End If
        If (colCollection Is Nothing) Then Return (0)
        With colCollection
            Dim colHandle As Collection
            Dim intDepth As Integer = 1
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                RaiseEvent collectionDepthEvent(Me.Name, .Count, intIndex1, intLevel)
                Try
                    colHandle = CType(.Item(intIndex1), Collection)
                Catch : End Try
                If Not (colHandle Is Nothing) Then
                    intDepth = Math.Max(intDepth, _
                                        collectionDepth(colHandle, intLevel + 1) + 1)
                End If
            Next intIndex1
            Return (intDepth)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Search collection for an object (scalar or collection)
    '
    '
    Public Function collectionFind(ByVal colCollection1 As Collection, _
                                   ByVal objTarget As Object) As Object
        If Not checkUsable_("collectionFind", "Returning 0") Then Return (0)
        If _OBJutilities.hasReferenceType(objTarget) _
           AndAlso _
           Not (TypeOf objTarget Is Collection) Then
            errorHandler_("Target " & _
                          _OBJutilities.object2String(objTarget) & " " & _
                          "is neither a scalar nor a collection", _
                          "collectionFind", _
                          "Returning 0")
            Return (0)
        End If
        Return (collectionFind_(colCollection1, objTarget, ""))
    End Function

    ' ----------------------------------------------------------------------
    ' Recursion on behalf of collectionFind
    '
    '
    Private Function collectionFind_(ByVal colCollection1 As Collection, _
                                     ByVal objTarget As Object, _
                                     ByVal strDewey As String) As Object
        With colCollection1
            Dim booMatch As Boolean
            Dim colItem As Collection
            Dim colTarget As Collection
            Dim intIndex1 As Integer
            Dim intLevel As Integer = _OBJutilities.items(strDewey, ".", False)
            Dim strItem As String
            Dim strMatch As String
            Dim strTarget As String
            If (TypeOf objTarget Is Collection) Then
                colTarget = CType(objTarget, Collection)
            End If
            For intIndex1 = 1 To .Count
                RaiseEvent collectionFindEvent(Me.Name, .Count, intIndex1, intLevel)
                If (TypeOf .Item(intIndex1) Is Collection) Then
                    colItem = CType(.Item(intIndex1), Collection)
                    If Not (colTarget Is Nothing) Then
                        ' Compare collection, to collection
                        If Me.collectionCompare(colItem, colTarget) Then
                            strMatch = CStr(intIndex1) : Exit For
                        End If
                    Else
                        ' Recursively search subcollection for target
                        strMatch = CStr(collectionFind_(colItem, _
                                                        objTarget, _
                                                        strDewey))
                        If strMatch <> "0" Then
                            Exit For
                        End If
                    End If
                Else
                    If matchObjects_(objTarget, .Item(intIndex1)) Then
                        strMatch = CStr(intIndex1)
                        Exit For
                    End If
                End If
            Next intIndex1
            If intIndex1 > .Count Then Return (0)
            If strDewey = "" Then Return (CInt(strMatch))
            Return (strDewey & "." & strMatch)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Form intersection set
    '
    '
    Public Function collectionIntersection(ByRef colCollection1 As Collection, _
                                           ByRef colCollection2 As Collection) _
                    As Collection
        If Not checkUsable_("collectionIntersection", _
                            "No intersection formed") Then Return (Nothing)
        Return collectionIntersection_(colCollection1, colCollection2, 0)
    End Function

    ' ----------------------------------------------------------------------
    ' Recursion on behalf of collectionIntersection
    '
    '
    Private Function collectionIntersection_(ByVal colCollection1 As Collection, _
                                             ByVal colCollection2 As Collection, _
                                             ByVal intLevel As Integer) _
                     As Collection
        Dim colIntersection As Collection = Me.list2Collection()
        If (colIntersection Is Nothing) Then Return (Nothing)
        Dim intIndex1 As Integer
        With colCollection2
            For intIndex1 = 1 To .Count
                If (TypeOf .Item(intIndex1) Is Collection) Then
                    Dim colInnerIntersection As Collection = _
                    collectionIntersection_(CType(.Item(intIndex1), Collection), _
                                            colCollection2, _
                                            intLevel + 1)
                    If Not Me.collectionAppend(colIntersection, colInnerIntersection) Then
                        Me.collectionClear(colIntersection)
                        Return (Nothing)
                    End If
                Else
                    If CStr(Me.collectionFind(colCollection1, .Item(intIndex1))) <> "0" Then
                        Try
                            colIntersection.Add(.Item(intIndex1))
                        Catch
                            errorHandler_("Unable to extend intersection: " & _
                                          Err.Number & " " & Err.Number, _
                                          "collectionComplement_", _
                                          "Returning partial result")
                        End Try
                    End If
                End If
            Next intIndex1
            Return (colIntersection)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Modify the collection item
    '
    '
    Public Function collectionItemSet(ByRef colCollection As Collection, _
                                      ByVal objIndex As Object, _
                                      ByVal objNewValue As Object) As Collection
        Dim colParent As Collection
        Dim intIndexInParent As Integer
        Dim objOldValue As Object = Me.dewey2Member(colCollection, _
                                                    objIndex, _
                                                    colParent, _
                                                    intIndexInParent)
        Dim strIndex As String = _OBJutilities.object2String(objIndex)
        RaiseEvent collectionItemSetEvent(Me.Name, _
                                          strIndex, _
                                          objOldValue, _
                                          objNewValue)
        If (colParent Is Nothing) Then
            errorHandler_("Cannot find member " & strIndex, _
                          "collectionItemSet", _
                          "Returning Nothing")
            Return (Nothing)
        End If
        Try
            With colParent
                .Remove(intIndexInParent)
                .Add(objNewValue, , intIndexInParent)
            End With
        Catch
            errorHandler_("Cannot replace member " & strIndex & ": " & _
                          Err.Number & " " & Err.Description, _
                          "collectionItemSet", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        Return (colCollection)
    End Function

    ' ----------------------------------------------------------------------
    ' Compares our collection string formats
    '
    '
    ' Disregarding minor format details, this method returns True when
    ' strCollection1 specifies (in the notation produced by collection2String)
    ' the identical collection (using the rules of collectionCompare) that's
    ' specified by strCollection2.
    '
    '
    Public Function collectionStringsCompare(ByVal strCollection1 As String, _
                                             ByVal strCollection2 As String) _
           As Boolean
        If Not checkUsable_("collectionStringsCompare", "Not comparing: returning False") Then
            Return (False)
        End If
        Dim colTest1 As Collection = string2Collection(strCollection1, booInspect:=False)
        Dim colTest2 As Collection = string2Collection(strCollection2, booInspect:=False)
        Return (Not (colTest1 Is Nothing) _
               AndAlso _
               Not (colTest2 Is Nothing) _
               AndAlso _
               collectionCompare(colTest1, colTest2))
    End Function

    ' ----------------------------------------------------------------------
    ' Collection to type collection
    '
    '
    ' This method analyzes a collection, and returns a new, separate
    ' collection. The new collection contains the string format, .Net
    ' name of each distinct type found in the input collection.
    '
    ' The names are in system format such as System.Int16 and there are
    ' no duplicates. The collection is keyed such that the key of each
    ' entry is identical to its data prefixed by an asterisk.
    '
    ' By default, and in "recursive" collections that contain collections
    ' as members, the name of the Collection type is included. However, if
    ' the optional booDrillDown parameter is present and True, then these
    ' collections are expanded. The name of the Collection type will not
    ' appear in the type collection, instead it will contain all types 
    ' found in all subcollections.
    '
    '
    Public Function collectionTypes(ByVal colCollection As Collection, _
                                    Optional ByVal booDrillDown As Boolean = False) _
           As Collection
        If Not checkUsable_("collectionTypes", "Returning Nothing") Then
            Return (Nothing)
        End If
        If (colCollection Is Nothing) Then Return (Nothing)
        Dim colTypes As Collection
        Try
            colTypes = New Collection
        Catch
            errorHandler_("Cannot create type collection: " & _
                          Err.Number & " " & Err.Description, _
                          "collectionTypes", _
                          "Returning Nothing", _
                          True)
            Return (Nothing)
        End Try
        If Not collectionTypes_(colCollection, booDrillDown, colTypes, 0) Then
            collectionClear(colTypes)
        End If
        Return (colTypes)
    End Function

    ' -----------------------------------------------------------------
    ' Recursion, on behalf of collectionTypes_
    '
    '
    Private Function collectionTypes_(ByVal colCollection As Collection, _
                                      ByVal booDrillDown As Boolean, _
                                      ByRef colTypes As Collection, _
                                      ByVal intLevel As Integer) As Boolean
        With colCollection
            Dim intIndex1 As Integer
            Dim strType As String
            For intIndex1 = 1 To .Count
                RaiseEvent collectionTypeEvent(Me.Name, .Count, intIndex1, intLevel)
                If (TypeOf .Item(intIndex1) Is Collection) AndAlso booDrillDown Then
                    If Not collectionTypes_(CType(.Item(intIndex1), Collection), _
                                            True, _
                                            colTypes, _
                                            intLevel + 1) Then
                        Exit For
                    End If
                Else
                    Try
                        strType = .Item(intIndex1).GetType.ToString
                        Dim strItem As String = CStr(colTypes.Item(string2Key_(strType)))
                    Catch
                        Try
                            colTypes.Add(strType, string2Key_(strType))
                        Catch
                            errorHandler_("Cannot extend type collection: " & _
                                          Err.Number & " " & Err.Description, _
                                          "collectionTypes", _
                                          "Returning False", _
                                          True)
                            Exit For
                        End Try
                    End Try
                End If
            Next intIndex1
            Return (intIndex1 > .Count)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Form union set
    '
    '
    Public Function collectionUnion(ByRef colCollection1 As Collection, _
                                    ByRef colCollection2 As Collection) _
                    As Collection
        If Not checkUsable_("collectionUnion", _
                            "No union formed") Then Return (Nothing)
        Dim colUnion As Collection
        Try
            colUnion = Me.list2Collection()
            Me.collectionAppend(colUnion, colCollection1)
            Me.collectionAppend(colUnion, colCollection2)
            Me.collection2Set(colUnion)
        Catch
            Me.collectionClear(colUnion)
        End Try
        Return (colUnion)
    End Function

    ' -----------------------------------------------------------------
    ' Set and return compare mode
    '
    '
    Public Property CompareAllowsClone() As Boolean
        Get
            If Not checkUsable_("CompareAllowsClone Get", "Returning False") Then
                Return (False)
            End If
            Return (USRstate.booCompareAllowsClone)
        End Get
        Set(ByVal booNewValue As Boolean)
            If Not checkUsable_("CompareAllowsClone Set", "No change made") Then
                Return
            End If
            USRstate.booCompareAllowsClone = booNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return the member identified by an index or by a Dewey number
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ---------------------------------------
    ' 07 16 03   Nilges         Added colParent and intItemIndex
    '
    '
    ' --- Basic syntax
    Public Overloads Function dewey2Member(ByVal colCollection As Collection, _
                                           ByVal objIndex As Object) As Object
        Dim colParent As Collection
        Dim intItemIndex As Integer
        Return (dewey2Member(colCollection, objIndex, colParent, intItemIndex))
    End Function
    ' --- Return parent and index
    Public Overloads Function dewey2Member(ByVal colCollection As Collection, _
                                           ByVal objIndex As Object, _
                                           ByRef colParent As Collection, _
                                           ByRef intItemIndex As Integer) As Object
        If Not checkUsable_("dewey2Member", "Returning Nothing") Then
            Return (Nothing)
        End If
        colParent = Nothing : intItemIndex = 0
        Dim strIndex As String
        Try
            strIndex = CStr(objIndex)
        Catch
            errorHandler_("Invalid index " & _OBJutilities.object2String(objIndex), _
                          "dewey2Member", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        Return (dewey2Member_(colCollection, strIndex, colParent, intItemIndex))
    End Function

    ' -----------------------------------------------------------------
    ' Recursion on behalf of dewey2Member_
    '
    '
    Private Function dewey2Member_(ByVal colCollection As Collection, _
                                   ByVal strIndex As String, _
                                   ByRef colParent As Collection, _
                                   ByRef intItemIndex As Integer) As Object
        Dim intIndex1 As Integer
        Dim strCurrentIndex As String = _OBJutilities.item(strIndex, 1, ".", False)
        Try
            intIndex1 = CInt(strCurrentIndex)
        Catch
            errorHandler_("Invalid index " & _OBJutilities.enquote(strCurrentIndex), _
                          "dewey2Member_", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        With colCollection
            If intIndex1 < 1 OrElse intIndex1 > .Count Then
                errorHandler_("Out-of-bounds index " & intIndex1, _
                              "dewey2Member_", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            Dim objElement As Object = .Item(intIndex1)
            If (TypeOf objElement Is Collection) _
               AndAlso _
               _OBJutilities.items(strIndex, ".", False) >= 2 Then
                Return (dewey2Member_(CType(objElement, Collection), _
                                     _OBJutilities.itemPhrase(strIndex, 2, -1, ".", False), _
                                     colParent, _
                                     intItemIndex))
            End If
            colParent = colCollection
            intItemIndex = intIndex1
            Return (objElement)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Pro-forma dispose
    '
    '
    Public Function dispose() As Boolean
        Me.mkUnusable()
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Extends the dictionary-structured index
    '
    '
    Public Function extendIndex(ByRef colCollection As Collection, _
                                ByVal objKey As Object, _
                                ByVal objData As Object) As Boolean
        If Not checkUsable_("extendIndex", _
                            "Did not change collection: returning False") Then
            Return (False)
        End If
        With colCollection
            Dim strKey As String
            Try
                strKey = CStr(objKey)
            Catch
                errorHandler_("Proposed key " & _
                              _OBJutilities.object2String(strKey) & " " & _
                              "cannot be converted to a string", _
                              "extendIndex", _
                              "Did not change collection: " & _
                              "returning False")
                Return (False)
            End Try
            strKey = string2Key_(strKey)
            Dim colEntry As Collection
            Try
                colEntry = New Collection
                With colEntry
                    .Add(strKey) : .Add(objData)
                End With
                colCollection.Add(colEntry, strKey)
            Catch
                errorHandler_("Cannot add entry: " & _
                              Err.Number & " " & Err.Description, _
                              "extendIndex", _
                              "Returning False")
                Return (False)
            End Try
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Tells whether collection is an index
    '
    '
    Public Overloads Function isIndex(ByVal colCollection As Collection) _
           As Boolean
        Dim strExplanation As String
        Return (Me.isIndex(colCollection, strExplanation))
    End Function
    Public Overloads Function isIndex(ByVal colCollection As Collection, _
                                      ByRef strExplanation As String) _
           As Boolean
        strExplanation = "Unknown error"
        If Not checkUsable_("isIndex", "Returning False") Then
            strExplanation = "Object instance is unusable"
            Return (False)
        End If
        If colCollection Is Nothing Then
            strExplanation = "Collection is Nothing"
            Return (False)
        End If
        With colCollection
            If .Count = 0 Then
                strExplanation = "Collection is empty"
                Return (False)
            End If
            Dim colHandle(1) As Collection
            Dim intIndex1 As Integer
            Dim strKey As String
            For intIndex1 = 1 To .Count
                Try
                    colHandle(0) = CType(.Item(intIndex1), Collection)
                Catch
                    strExplanation = "Cannot convert item " & intIndex1 & " " & _
                                     "to a collection"
                    Return (False)
                End Try
                With colHandle(0)
                    If .Count <> 2 Then
                        strExplanation = "Item " & intIndex1 & " " & _
                                         "doesn't contain two entries"
                        Return (False)
                    End If
                    Try
                        strKey = "(unknown)"
                        strKey = CStr(.Item(1))
                        colHandle(1) = CType(colCollection.Item(strKey), _
                                             Collection)
                    Catch
                        strExplanation = "Key at index " & intIndex1 & " " & _
                                         _OBJutilities.enquote(strKey) & " " & _
                                         "cannot be used"
                        Return (False)
                    End Try
                    If Not (colHandle(0) Is colHandle(1)) Then
                        strExplanation = "Key at index " & intIndex1 & " " & _
                                         _OBJutilities.enquote(strKey) & " " & _
                                         "doesn't access its own entry"
                        Return (False)
                    End If
                End With
            Next intIndex1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' List to collection
    '
    '
    Public Function list2Collection(ByVal ParamArray objItems() As Object) _
           As Collection
        If Not checkUsable_("list2Collection", "Returning Nothing") Then
            Return (Nothing)
        End If
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch
            errorHandler_("Cannot create collection: " & _
                          Err.Number & " " & Err.Description, _
                          "list2Collection", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        Dim intIndex1 As Integer
        For intIndex1 = 0 To UBound(objItems)
            Try
                colNew.Add(objItems(intIndex1))
            Catch
                errorHandler_("Cannot create collection: " & _
                              Err.Number & " " & Err.Description, _
                              "list2Collection", _
                              "Returning partial result collection")
                Exit For
            End Try
        Next intIndex1
        Return (colNew)
    End Function

    ' -----------------------------------------------------------------
    ' Make one random collection
    '
    '
    ' --- Defaults
    Public Overloads Function mkRandomCollection() As Collection
        Return (mkRandomCollection(DEFAULT_RANDOM_MIN, _
                                  DEFAULT_RANDOM_MAX, _
                                  DEFAULT_RANDOM_TYPES, _
                                  DEFAULT_RANDOM_DEPTH))
    End Function
    ' --- Minimum and maximum items
    Public Overloads Function mkRandomCollection(ByVal intMinItems As Integer, _
                                                 ByVal intMaxItems As Integer) As Collection
        Return (mkRandomCollection(intMinItems, _
                                  intMaxItems, _
                                  DEFAULT_RANDOM_TYPES, _
                                  DEFAULT_RANDOM_DEPTH))
    End Function
    ' --- Minimum and maximum items: list of types
    Public Overloads Function mkRandomCollection(ByVal intMinItems As Integer, _
                                                 ByVal intMaxItems As Integer, _
                                                 ByVal strTypes As String) As Collection
        Return (mkRandomCollection(intMinItems, _
                                  intMaxItems, _
                                  strTypes, _
                                  DEFAULT_RANDOM_DEPTH))
    End Function
    ' --- Minimum and maximum items: list of types: max recursion depth
    Public Overloads Function mkRandomCollection(ByVal intMinItems As Integer, _
                                                 ByVal intMaxItems As Integer, _
                                                 ByVal strTypes As String, _
                                                 ByVal intMaxDepth As Integer) _
           As Collection
        Return (mkRandomCollection_(intMinItems, _
                                   intMaxItems, _
                                   strTypes, _
                                   intMaxDepth, _
                                   0))
    End Function
    ' --- Private recursion logic
    Private Overloads Function mkRandomCollection_(ByVal intMinItems As Integer, _
                                                   ByVal intMaxItems As Integer, _
                                                   ByVal strTypes As String, _
                                                   ByVal intMaxDepth As Integer, _
                                                   ByVal intRecursionDepth As Integer) _
           As Collection
        Dim colRandom As Collection
        Try
            colRandom = New Collection
        Catch
            errorHandler_("Can't create random collection: " & _
                          Err.Number & " " & Err.Description, _
                          "mkRandomCollection", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        Dim intCount As Integer
        Dim intEntries As Integer = CInt(Rnd() * intMaxItems)
        Dim objEntry As Object
        For intCount = 1 To intEntries
            RaiseEvent mkRandomCollectionEvent(Me.Name, intEntries, intCount, intRecursionDepth)
            Select Case UCase(_OBJutilities.word(strTypes, _
                                                 Math.Max(1, CInt(Rnd() * _OBJutilities.words(strTypes)))))
                Case "BOOLEAN" : objEntry = (Rnd() > 0.5)
                Case "BYTE" : objEntry = CByte(Rnd() * 255)
                Case "SHORT" : objEntry = CShort(Rnd() * 10000 * CShort(IIf(Rnd() < 0.5, -1, 1)))
                Case "INTEGER" : objEntry = CInt(Rnd() * 1000000 * CInt(IIf(Rnd() < 0.5, -1, 1)))
                Case "LONG" : objEntry = CLng(Rnd() * 1000000000000000 * CLng(IIf(Rnd() < 0.5, -1, 1)))
                Case "SINGLE" : objEntry = CSng(Math.Round(Rnd, CInt(Rnd() * 10)))
                Case "DOUBLE" : objEntry = CDbl(Math.Round(Rnd, CInt(Rnd() * 10)))
                Case "STRING" : objEntry = mkRandomCollection_randomString_()
                Case "COLLECTION"
                    If intRecursionDepth >= intMaxDepth Then
                        objEntry = "Can't provide collection, recursion depth exceeded"
                    Else
                        objEntry = mkRandomCollection_ _
                                   (intMinItems, intMaxItems, strTypes, intMaxDepth, _
                                    intRecursionDepth + 1)
                    End If
                Case "STRINGBUILDER"
                    Try
                        objEntry = New System.Text.StringBuilder
                    Catch
                        errorHandler_("Could not create stringBuilder: " & _
                                      Err.Number & " " & Err.Description, _
                                      "mkRandomCollection", _
                                      "Destroying random collection and returning Nothing", _
                                      True)
                        Me.collectionClear(colRandom)
                        Return (Nothing)
                    End Try
                Case Else
                    errorHandler_("Unexpected case", _
                                  "mkRandomCollection", _
                                  "Destroying random collection and returning Nothing", _
                                  True)
                    Me.collectionClear(colRandom)
                    Return (Nothing)
            End Select
            Try
                colRandom.Add(objEntry)
            Catch
                errorHandler_("Cannot extend random collection: " & _
                              Err.Number & " " & Err.Description, _
                              "mkRandomCollection", _
                              "Destroying random collection and returning Nothing", _
                              True)
                Me.collectionClear(colRandom)
                Return (Nothing)
            End Try
        Next intCount
        Return (colRandom)
    End Function

    ' -----------------------------------------------------------------
    ' Return a random string on behalf of mkRandomCollection
    '
    '
    Private Function mkRandomCollection_randomString_() As String
        Dim intLength As Integer
        Dim strRandom As String
        For intLength = 0 To CInt(Rnd() * 32)
            strRandom &= Mid("abcdef", CInt(Rnd() * 5) + 1, 1)
        Next intLength
        Return (strRandom)
    End Function

    ' ----------------------------------------------------------------------
    ' Mark the object as not usable
    '
    '
    Public Function mkUnusable() As Boolean
        USRstate.booUsable = False
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Return and set an instance name
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

    ' ----------------------------------------------------------------------
    ' Retrieve item from dictionary
    '
    '
    Public Function retrieveFromIndex(ByVal colIndex As Collection, _
                                      ByVal objKey As Object) As Boolean
        If Not checkUsable_("retrieveFromIndex", _
                            "Returning Nothing") Then Return (Nothing)
        With colIndex
            Dim strKey As String
            Try
                strKey = CStr(objKey)
            Catch
                errorHandler_("Cannot convert key " & _
                              _OBJutilities.object2String(objKey) & " " & _
                              "to a string: " & _
                              Err.Number & " " & Err.Description, _
                              "retrieveFromIndex", _
                              "Returning Nothing")
                Return (Nothing)
            End Try
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Convert string returned by collection2String back to a collection
    '
    '
    ' This method converts a string that has been created by the
    ' collection2String method back into a new collection, where possible,
    ' and this method returns this new collection as a function value.
    '
    ' The string input to this method must be in the format created by
    ' calling collection2String with booReadable:=False because booReadable
    ' is the optional parameter, determining whether a readable or less-
    ' readable string is returned. This method cannot convert readable
    ' strings back to collections, and it can only convert the "palindromic"
    ' (reversible) subset of less-readable, strongly typed strings.
    '
    ' See the documentation of collection2String for a discussion of what we
    ' mean by palindromes and readable collections.
    '
    ' Where the string consists of a comma-delimited list of items, they
    ' are converted according to the following rules. If all items can be
    ' converted using rules 1..3 below, then the output collection is an
    ' "ok" collection, known to accurately represent the original collection
    ' from which the input string was built. In this case, the input string
    ' is a palindrome as explained in the collection2String documentation.
    '
    '
    '      1. If the entire list is noCollection (representing a
    '         collection that is unallocated) this method will return
    '         Nothing.
    '
    '      2. If the entire list is emptyCollection (representing a
    '         collection that is allocated but has no members) this 
    '         method will return a new collection with no members.
    '
    '      3. Each value that is in the format canonicalType(value) is
    '         converted using the rules of the string2Object method to
    '         a value of one of the "canonical", or standard types:
    '         Boolean, Byte, Short, Integer, Long, Single, Double, or 
    '         String.
    '
    '         If the canonicalType is String then value must be in
    '         quotes and may contain XML equivalents for non-graphic
    '         characters, quote, comma, and parentheses.
    '
    '      4. Each value that is a parenthesized list is converted 
    '         recursively into a collection using this method.
    '
    '      5. The unquoted string Nothing results in a collection entry containing
    '         Nothing
    '
    '      6. All other comma-delimited items are converted back to an
    '         object by object2String: application of this rule means that
    '         the returned collection is not fully "OK" and probably is not
    '         a clone of the original object.
    '
    '
    ' The basic syntax of this method is string2Collection(s), returning a
    ' new collection c (or Nothing on failure).  
    '
    ' The overloaded syntax string2Collection(s, OK) will set OK to -1 or 
    ' to 0 to indicate a fully valid conversion; -1 indicates a valid
    ' conversion, 0 indicates that for one or more items, rule 6 above
    ' was used and that as a result, the returned collection probably does
    ' not clone the original collection, used to create the input string.
    '
    ' By default and only when all items in the input string were converted
    ' by rules 1..5 above, this method will check the result by ensuring that 
    ' the output collection converts to a string (using collection2String) that is
    ' identical, as a collection string, to the input string. This check can
    ' be suppressed using the optional parameter booInspect:=False. It won't
    ' be made if not all items in the input string are type(value), (collection),
    ' or Nothing.
    ' 
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 01 07 03   Nilges         Version 1.0
    ' 02 17 03   Nilges         Added strDelim and completed code
    ' 06 26 03   Nilges         1. Changed the delimiter to a comma with no
    '                              trailing blanks
    '                           2. Bug: misplaced parenthesis
    '                           3. Bug: Parentheses not balanced
    ' 07 03 03   Nilges         1. Added booInspect
    ' 07 04 03   Nilges         1. Require object2String syntax
    '                           2. Changed type of the OK byref to integer
    '                           3. Removed strDelim
    '
    '
    ' --- Simple call
    Public Overloads Function string2Collection(ByVal strInstring As String, _
                                                Optional ByVal booInspect As Boolean = True) _
           As Collection
        Dim intOK As Integer
        Return string2Collection(strInstring, intOK, booInspect:=False)
    End Function
    ' --- Call with OK set
    Public Overloads Function string2Collection(ByVal strInstring As String, _
                                                ByRef intOK As Integer, _
                                                Optional ByVal booInspect As Boolean = True) _
           As Collection
        Return (string2Collection_(strInstring, intOK, booInspect, 0))
    End Function
    ' --- Common logic
    Private Function string2Collection_(ByVal strInstring As String, _
                                        ByRef intOK As Integer, _
                                        ByVal booInspect As Boolean, _
                                        ByVal intLevel As Integer) _
           As Collection
        If Not checkUsable_("string2Collection", "Returning Nothing") Then
            Return (Nothing)
        End If
        intOK = 0
        If UCase(strInstring) = "NOCOLLECTION" Then
            intOK = -1 : Return (Nothing)
        End If
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch
            errorHandler_("Cannot create output collection object", _
                          "string2Collection", _
                          Err.Number & " " & Err.Description & _
                          vbNewLine & vbNewLine & _
                          "Returning Nothing: booOK is False", _
                          True)
            Return (Nothing)
        End Try
        If UCase(strInstring) = "EMPTYCOLLECTION" Then
            intOK = -1 : Return (colNew)
        End If
        Dim colScanned As Collection = string2Collection_scanner_(strInstring)
        If (colScanned Is Nothing) Then
            errorHandler_("Cannot scan the input string", _
                          "string2Collection", _
                          "Returning empty collection: booOK is False")
            Return (colNew)
        End If
        Dim booNest As Boolean
        Dim colEntry As Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim objNext As Object
        Dim booOKwork As Boolean = True
        Dim dblNext As Double
        Dim strNext As String
        With colScanned
            For intIndex1 = 1 To .Count
                RaiseEvent string2CollectionEvent(Me.Name, .Count, intIndex1, intLevel)
                colEntry = CType(.Item(intIndex1), Collection)
                With colEntry
                    booNest = CBool(.Item(1))
                    strNext = CStr(.Item(2))
                End With
                objNext = Nothing
                If booNest Then
                    objNext = string2Collection_(strNext, intOK, booInspect, intLevel + 1)
                    booOKwork = booOKwork AndAlso (TypeOf objNext Is Collection)
                Else
                    Try
                        objNext = _OBJutilities.string2Object(strNext)
                    Catch : End Try
                    booOKwork = booOKwork _
                                AndAlso _
                                ((objNext Is Nothing) AndAlso UCase(strNext) = "NOTHING" _
                                 OrElse _
                                 Not (objNext Is Nothing) _
                                 AndAlso _
                                 (Not _OBJutilities.hasReferenceType(objNext)))
                End If
                Try
                    colNew.Add(objNext)
                Catch
                    errorHandler_("Cannot add object to collection", _
                                  "string2Collection", _
                                  Err.Number & " " & Err.Description & _
                                  vbNewLine & vbNewLine & _
                                  "Returning partial collection: intOK is 0 (False)", _
                                  True)
                    Return (colNew)
                End Try
            Next intIndex1
        End With
        intOK = CInt(IIf(booOKwork, -1, 0))
        If booOKwork _
           AndAlso _
           booInspect _
           AndAlso _
           Not collectionStringsCompare(strInstring, collection2String(colNew, booInspect:=False)) Then
            errorHandler_("Inspection failure", _
                          "string2Collection", _
                          "Returning collection which is probably invalid", _
                          True)
        End If
        Return (colNew)
    End Function

    ' -----------------------------------------------------------------
    ' Scan the input string on behalf of string2Collection
    '
    '
    ' For each comma-delimited item this method creates one entry in its 
    ' output collection colScanned.
    '
    ' Each comma-delimited item of colScanned is converted a two-, or one-item 
    ' subcollection in colScanned such that item(1) indicates the type of the
    ' entry and item(2), in most cases, is its string value.
    '
    '
    '      *  If the comma-delimited item is in the form type(value)
    '         or it is the string Nothing, then the subcollection item(1)
    '         is set to False and item(2) is a string in the format
    '         type(value) (or Nothing), acceptable to the string2Object method.
    '
    '      *  If the item is a parenthesized (collection), item(1) is
    '         True and item(2) is the serialized collection without 
    '         its parentheses.
    '
    '      *  If the comma-delimited item is anything else we raise an
    '         error and return Nothing
    '
    '
    Private Function string2Collection_scanner_(ByVal strInstring As String) _
            As Collection
        Dim colScanned As Collection
        Try
            colScanned = New Collection
        Catch
            errorHandler_("Cannot create the colScanned collection: " & _
                          Err.Number & " " & Err.Description, _
                          "string2Collection_scanner_", _
                          "Returning Nothing", _
                          True)
            Return (Nothing)
        End Try
        Dim booEntryType As Boolean
        Dim colEntry As Collection
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim strInstringTrim As String = Trim(strInstring)
        Dim strNext As String
        Do While intIndex1 <= Len(strInstringTrim)
            ' --- Skip blanks
            For intIndex1 = intIndex1 To Len(strInstringTrim)
                If Mid(strInstringTrim, intIndex1, 1) <> " " Then Exit For
            Next intIndex1
            If intIndex1 > Len(strInstringTrim) Then Exit Do
            ' --- Create new entry subcollection        
            Try
                colEntry = New Collection
            Catch ex As Exception
                errorHandler_("Cannot create scanned collection entry: " & _
                              Err.Number & " " & Err.Description, _
                              "string2Collection_scanner_", _
                              "Returning Nothing", _
                              True)
                Return (Nothing)
            End Try
            ' --- Test type of comma-delimited item        
            If Mid(strInstringTrim, intIndex1, 1) = "(" Then
                ' Contained collection       
                booEntryType = True
                intIndex2 = _OBJutilities.findBalParenthesis(strInstring & ")", intIndex1)
                intIndex1 += 1
                strNext = Trim(Mid(strInstringTrim, _
                                   intIndex1, _
                                   intIndex2 - intIndex1))
                intIndex2 += 1
                If intIndex2 < Len(strInstring) AndAlso Mid(strInstring, intIndex2, 1) <> "," Then
                    errorHandler_("Parenthesis is not followed by a comma", _
                                  "string2Collection_scanner_", _
                                  "Returning Nothing", _
                                  False)
                    Return (Nothing)
                End If
            Else
                ' Hopefully: type(value) or Nothing
                booEntryType = False
                intIndex2 = InStr(intIndex1, strInstring & ",", ",")
                strNext = Trim(Mid(strInstringTrim, _
                                   intIndex1, _
                                   intIndex2 - intIndex1))
                If UCase(strNext) <> "NOTHING" _
                   AndAlso _
                   (Mid(strNext, Len(strNext)) <> ")" _
                    OrElse _
                    InStr(strNext, "(") = 0) Then
                    errorHandler_("string2Collection item " & _
                                  _OBJutilities.enquote(strNext) & " " & _
                                  "is not a valid token for object2String " & _
                                  "in the format type(value)", _
                                  "", _
                                  "Returning a scan collection of Nothing")
                    Return (Nothing)
                End If
            End If
            Try
                With colEntry
                    .Add(booEntryType)
                    .Add(strNext)
                End With
                colScanned.Add(colEntry)
            Catch
                errorHandler_("Cannot extend the colScanned collection: " & _
                              Err.Number & " " & Err.Description, _
                              "string2Collection_scanner_", _
                              "Returning Nothing", _
                              True)
                Return (Nothing)
            End Try
            intIndex1 = intIndex2 + 1
        Loop
        Return (colScanned)
    End Function

    ' -----------------------------------------------------------------
    ' Test this class
    '
    '
    ' --- Default count
    Public Overloads Function test(ByRef strReport As String) As Boolean
        Return (test(strReport, STRESS_TEST_COUNT))
    End Function
    ' --- Specified count
    Public Overloads Function test(ByRef strReport As String, _
                                   ByVal intCount As Integer) As Boolean
        If intCount < 0 Then
            errorHandlerShared_("Invalid test count " & intCount, _
                                ClassName, _
                                "test", _
                                "No test conducted, returning False")
            Return (False)
        End If
        Dim intTestNumber As Integer
        Dim objResult As Object
        Dim strCalculusReport As String
        Dim strSerialized As String = "Can't get test result"
        strReport = "Testing " & Me.ClassName & " using instance " & Me.Name & " " & _
                    "at " & Now & _
                    vbNewLine & vbNewLine
        For intTestNumber = 1 To intCount
            RaiseEvent testEvent(Me.Name, intCount, intTestNumber)
            Try
                objResult = Me.collectionCalculusWithReport(STRESS_TEST_STEP, _
                                                            strCalculusReport)
            Catch
                errorHandler_("Error occured during the test: " & _
                              Err.Number & " " & Err.Description, _
                              "test", _
                              "Returning False", _
                              True)
                Return (False)
            End Try
            If Me.Usable AndAlso (TypeOf objResult Is Collection) Then
                strSerialized = Me.collection2String(CType(objResult, Collection))
            Else
                strSerialized = _OBJutilities.object2String(strSerialized, True)
            End If
            strReport &= vbNewLine & vbNewLine & _
                         "Test " & intTestNumber & " of " & intCount & _
                         vbNewLine & vbNewLine & _
                         strCalculusReport & _
                         vbNewLine & vbNewLine & _
                         "Test produced the object " & _
                         _OBJutilities.ellipsis(strSerialized, 64)
        Next intTestNumber
    End Function

    ' -----------------------------------------------------------------
    ' Return the test expression
    '
    '
    Private Shared ReadOnly Property TestExpression() As String
        Get
            Return (STRESS_TEST_STEP)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return the object's Usability status
    '
    '
    Public ReadOnly Property Usable() As Boolean
        Get
            Return (USRstate.booUsable)
        End Get
    End Property

#End Region ' Public procedures except collectionCalculus      

#Region " Private procedures "
    
    ' ***** Private Procedures ****************************************
    
    ' -----------------------------------------------------------------
    ' Check the usability of the object
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String, _
                                  ByVal strAction As String) As Boolean
        If Not Me.Usable Then
            errorHandler_("Object instance isn't usable", _
                          strProcedure, _
                          strAction)
        End If        
        Return(True)
    End Function                                  
    
    ' -----------------------------------------------------------------
    ' Error handling interface
    '
    '
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String)
        errorHandler_(strMessage, strProcedure, strHelp, False)
    End Sub                                       
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String, _
                                        ByVal booMkUnusable As Boolean)
        errorHandlerShared_(strMessage, _
                            Me.Name, _
                            strProcedure, _
                            strHelp & _
                            CStr(IIf(booMkUnusable, _
                                     vbNewline & vbNewline & _
                                     "The object has been marked unusable", _
                                     "")))
        If booMkUnusable Then Me.mkUnusable
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Shared error handling
    '
    '
    Private Shared Sub errorHandlerShared_(ByVal strMessage As String, _
                                           ByVal strName As String, _
                                           ByVal strProcedure As String, _
                                           ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, strName, strProcedure, strHelp)
    End Sub   
    
    ' -----------------------------------------------------------------
    ' Match objects taking CompareAllowsClone into consideration
    '
    '
    Private Function matchObjects_(ByVal objObject1 As Object, _
                                   ByVal objObject2 As Object) As Boolean
        Dim booRefType1 As Boolean = _OBJutilities.hasReferenceType(objObject1)                                  
        Dim booRefType2 As Boolean = _OBJutilities.hasReferenceType(objObject2)                                  
        If Not booRefType1 AndAlso Not booRefType2 Then
            ' Scalar to scalar comparision using strings
            Return(CStr(objObject1) = CStr(objObject2))
        ElseIf booRefType1 And booRefType2 Then    
            ' Object to object comparision
            If Me.CompareAllowsClone Then
                ' Return True only if collections passed that are clones
                If Not (TypeOf objObject1 Is Collection) Then Return(False)
                If Not (TypeOf objObject2 Is Collection) Then Return(False)
                Return(Me.collectionCompare(CType(objObject1, Collection), _
                                            CType(objObject2, Collection)))
            Else
                ' Use ISwise compare
                Return (objObject1 Is objObject2)                                            
            End If                                
        End If           
        Return(False)
    End Function                
    
    ' ----------------------------------------------------------------------
    ' Object to scalar value, type(value) or collection2String format
    '
    '
    Private Function object2Display_(ByVal objObject As Object) As String
        If (objObject Is Nothing) Then
            Return("Nothing")
        ElseIf (TypeOf objObject Is Collection) Then
            Return(Me.collection2String(CType(objObject, Collection), _
                                        booReadable:=True, _
                                        booInspect:=False))
        ElseIf _OBJutilities.hasReferenceType(objObject) Then
            Return(CStr(objObject))   
        Else
            Return(_OBJutilities.object2String(objObject))                                                                                     
        End If     
    End Function  
    
    ' ----------------------------------------------------------------------
    ' Convert string to key
    '
    '
    Private Function string2Key_(ByVal strKey As String) As String
        If Not IsNumeric(strKey) Then Return(strKey)
        Return("*" & strKey)
    End Function                                                                                          

#End Region                               
                                                                            
#Region " collectionCalculus "
    
    ' ----------------------------------------------------------------------
    ' Collection macro processor
    '
    '
    Public Function collectionCalculus(ByVal strExpression As String, _
                                       ParamArray objCollectionList() As Object) _
           As Object
        If Not checkUsable_("collectionCalculus", "Returning Nothing") Then
            Return(Nothing)
        End If        
        Dim strReport As String
        Select Case UBound(objCollectionList)
            Case -1:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport))
            Case 1:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1)))
            Case 3:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3)))
            Case 5:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5)))
            Case 7:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5), _
                                                     objCollectionList(6), objCollectionList(7)))
            Case 9:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     False, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5), _
                                                     objCollectionList(6), objCollectionList(7), _
                                                     objCollectionList(8), objCollectionList(9)))
            Case Else:
                errorHandler_("Odd number of collections in list, " & _
                              "or more than allowed", _
                              "collectionCalculus", _
                              "Returning Nothing")
                Return(Nothing)                                                                                  
        End Select        
    End Function       
    
    ' ----------------------------------------------------------------------
    ' Collection calculus with a nice report
    '
    '
    Public Function collectionCalculusWithReport(ByVal strExpression As String, _
                                                 ByRef strReport As String, _
                                                 ParamArray objCollectionList() As Object) _
           As Object   
        Select Case UBound(objCollectionList)
            Case -1:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport))
            Case 1:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1)))
            Case 3:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3)))
            Case 5:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5)))
            Case 7:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5), _
                                                     objCollectionList(6), objCollectionList(7)))
            Case 9:
                Return(collectionCalculusWithReport_(strExpression, _
                                                     True, _
                                                     strReport, _
                                                     objCollectionList(0), objCollectionList(1), _
                                                     objCollectionList(2), objCollectionList(3), _
                                                     objCollectionList(4), objCollectionList(5), _
                                                     objCollectionList(6), objCollectionList(7), _
                                                     objCollectionList(8), objCollectionList(9)))
            Case Else:
                errorHandler_("Odd number of collections in list, " & _
                              "or more than allowed", _
                              "collectionCalculusWithReport", _
                              "Returning Nothing")
                Return(Nothing)                                                                                  
        End Select        
    End Function           
    
    ' ----------------------------------------------------------------------
    ' Common collection calculus logic
    '
    '
    Private Function collectionCalculusWithReport_(ByVal strExpression As String, _
                                                   ByVal booReport As Boolean, _
                                                   ByRef strReport As String, _
                                                   ParamArray objCollectionList() As Object) _
           As Object   
        If Not checkUsable_("collectionCalculusWithReport", "Returning Nothing") Then
            Return(Nothing)
        End If
        Dim objValue As Object 
        If booReport Then strReport = ""
        Try             
            ' --- Create the byref collection list
            If UBound(objCollectionList) Mod 2 = 0 Then
                collectionCalculusWithReport_errorHandler_("initialization", _
                                                            "Incorrect parameter list size", _
                                                            "Returning Nothing")
                Return(Nothing)                          
            End If        
            Dim colByRef As Collection
            Try
                colByRef = New Collection
            Catch  
                collectionCalculusWithReport_errorHandler_("initialization", _
                                                           "Cannot create the ByRef collection list: " & _
                                                           Err.Number & " " & Err.Description, _
                                                           "Returning Nothing")
            End Try        
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim strName As String
            For intIndex1 = 0 To UBound(objCollectionList) Step 2
                intIndex2 = intIndex1
                If Not (TypeOf objCollectionList(intIndex1) Is System.String) _
                OrElse _
                Not (TypeOf objCollectionList(intIndex2) Is Collection) Then
                    collectionCalculusWithReport_errorHandler_("initialization", _
                                                                "ByRef collection list contains invalid types", _
                                                                "Returning Nothing")
                    Return(Nothing)                          
                End If  
                Try
                    strName = CStr(objCollectionList(intIndex1))
                    If Mid(strName, 1, 1) <> "_" Then
                        collectionCalculusWithReport_errorHandler_("initialization", _
                                                                   "The collection name " & _
                                                                   _OBJutilities.enquote(strName) & " " & _
                                                                   "does not start with an underscore", _
                                                                   "Raising error")
                    End If                
                    Dim colEntry As Collection = New Collection
                    With colEntry
                        .Add(objCollectionList(intIndex2)): .Add(strName)
                    End With                    
                    colByRef.Add(colEntry, strName)
                Catch  
                    collectionCalculusWithReport_errorHandler_("initialization", _
                                                               "Can't expand ByRef collection list: " & _
                                                               Err.Number & " " & Err.Description, _
                                                               "Returning Nothing")
                    Return(Nothing)                          
                End Try                       
            Next intIndex1        
            ' --- Parse the expresion 
            ' Lexical analysis      
            Dim colScanned As Collection = _
                collectionCalculusWithReport_expressionScan_(strExpression, strReport)
            If (colScanned Is Nothing) Then
                collectionCalculusWithReport_errorHandler_("scan", _
                                                           "Not able to scan the calculus expression", _
                                                           "Returning Nothing")
            End If    
            ' Parsing                
            Dim colParsed As Collection = _
                collectionCalculusWithReport_expressionParse_(colScanned, strReport)
            If (colParsed Is Nothing) Then
                collectionCalculusWithReport_errorHandler_("scan", _
                                                           "Not able to parse the calculus expression", _
                                                           "Returning Nothing")
            End If                    
            ' --- Interpret the expression
            objValue = (collectionCalculusWithReport_interpreter__ _
                        (strExpression, colParsed, colByRef, strReport))
        Catch
            collectionCalculusWithReport_errorReport_(Err.Number & " " & Err.Description, strReport)
        End Try
        If booReport Then
            Dim strAnalytic As String
            If (Typeof objValue Is Collection) Then
                strAnalytic = Me.collection2String(CType(objValue, Collection))
            Else
                strAnalytic = _OBJutilities.object2String(objValue)                
            End If            
            _OBJutilities.append(strReport, _
                                 vbNewline & vbNewline, _
                                 "collectionCalculus result (readable): " & _
                                 _OBJutilities.ellipsis(object2Display_(objValue), 32) & _
                                 vbNewline & vbNewline & _
                                 "collectionCalculus result (analytic): " & _
                                 _OBJutilities.ellipsis(strAnalytic, 32))
        End If        
        Return(objValue)
    End Function      
    
    ' ----------------------------------------------------------------------
    ' Error handler for collectionCalculus
    '
    '
    Private Sub collectionCalculusWithReport_errorHandler_(ByVal strPhase As String, _
                                                           ByVal strError As String, _
                                                           ByVal strHelp As String)
        errorHandler_("Error in " & strPhase & " phase: " & strError, _
                      "collectionCalculusWithReport_errorHandler_", _
                      strHelp)
    End Sub           
    
    ' ----------------------------------------------------------------------
    ' Report an error in collectionCalculus
    '
    '
    Private Sub collectionCalculusWithReport_errorReport_(ByVal strError As String, _
                                                          ByRef strReport As String)
        If (strReport Is Nothing) Then
            ' --- Raise the error
            errorHandler_(strError, "collectionCalculusWithReport_errorReport_", "")
        Else
            ' --- Append the error information to the report
            _OBJutilities.append(strReport, _
                                 vbNewline & vbNewline, _
                                 strError)
        End If        
    End Sub                                                                                                           
    
    ' ----------------------------------------------------------------------
    ' Parse expression on behalf of the collectionCalculus methods
    '
    '
    ' This method uses the following BNF specification of the calculus:
    '
    '
    '   calculusExpression := functionCall [ : calculusExpression ]
    '   functionCall := methodName [ argumentList ]
    '   argumentList := ( [ argumentListBody ] )
    '   argumentListBody := arg [ , argumentListBody ]
    '   arg := functionCall | byRefName | OTHER_TEXT
    '   methodName := COLLECTIONCALCULUS_METHOD (one of the supported
    '                 method names) 
    '   byRefName := VB_IDENTIFIER_WITHUNDERSCORE
    '
    '
    ' It produces an unkeyed collection which is the Reverse Polish
    ' representaton of the parsed calculus expression. Each item in 
    ' this collection is a subcollection with 1 or 2 items:
    '
    '
    '      *  Item(1) is a method name, 0_PUSHLITERAL, or 1_PUSHBYREF
    '
    '      *  If Item(1) is a method, Item(2) is missing; if Item(1)
    '         is 0_PUSHLITERAL or 1_PUSHBYREF then Item(2) is a string
    '
    '
    Private Function collectionCalculusWithReport_expressionParse_ _
                     (ByVal colScanned As Collection, _
                      ByRef strReport As String) _
            As Collection
        Dim colParse As Collection
        Try
            colParse = New Collection
        Catch  
            errorHandler_("Cannot create the parse collection: " & _
                          Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport", _
                          "Returning Nothing", _
                          True)
            Return(Nothing)                          
        End Try        
        Dim intIndex1 As Integer = 1
        If Not collectionCalculusWithReport_expressionParse__calculusExpression_ _
               (colScanned, intIndex1, colParse, colScanned.Count + 1) _
           OrElse _
           intIndex1 <= colScanned.Count Then
            collectionCalculusWithReport_errorHandler_("parse", _
                                                       "Cannot parse the calculus expression", _
                                                       "")
            Me.collectionClear(colParse)                          
            Return(Nothing)                          
        End If     
        If Not (strReport Is Nothing) Then
            _OBJutilities.append(strReport, _
                                 vbNewline, _
                                 "Interpreter code: " & _
                                 Me.collection2String(colParse, _
                                                      booReadable:=True, _
                                                      booInspect:=False))        
        End If        
        Return(colParse)          
    End Function               
    
    ' ----------------------------------------------------------------------
    ' arg := functionCall | byRefName | OTHER_TEXT
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__arg_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "arg")                      
        If Not collectionCalculusWithReport_expressionParse__functionCall_ _
               (colScanned, intIndex, colParse, intEndIndex) _
           AndAlso _
           Not collectionCalculusWithReport_expressionParse__byRefName_ _
               (colScanned, intIndex, colParse, intEndIndex)  Then
            If CStr(CType(colScanned.Item(intIndex), Collection).Item(1)) = "text" Then                       
                collectionCalculusWithReport_expressionParse__codeGen_ _
                (colParse, _
                 "0_PUSHLITERAL", _
                 CStr(CType(colScanned.Item(intIndex), Collection).Item(2)))
                intIndex += 1    
            Else 
                Return(False)                
            End If         
        End If               
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' argumentList := ( [ argumentListBody ] )
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__argumentList_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "argumentList")   
        Dim intIndex1 As Integer = intIndex   
        If Not collectionCalculusWithReport_expressionParse__checkToken_ _
               (colScanned, intIndex, intEndIndex, "parenthesis", "(") Then
            Return(False)
        End If       
        Dim intIndex2 As Integer = _                           
            collectionCalculusWithReport_expressionParse__findBalParenthesis_ _
            (colScanned, intIndex, intEndIndex)
        Dim intFrameSize As Integer             
        If Not collectionCalculusWithReport_expressionParse__argumentListBody_ _
               (colScanned, intIndex, colParse, intIndex2, intFrameSize) Then
            If intIndex <> intIndex2 Then               
                collectionCalculusWithReport_expressionParse__errorHandler_ _
                ("Left parenthesis is not followed by an argument list body", _
                 "", _
                 intIndex1, _
                 colScanned)               
                intIndex = intIndex1               
                Return(False)
            End If                
        End If
        If Not collectionCalculusWithReport_expressionParse__checkToken_ _
               (colScanned, intIndex, intEndIndex, "parenthesis", ")") Then
            collectionCalculusWithReport_expressionParse__errorHandler_ _
            ("Argument list body is not followed by right parenthesis", _
             "", _
             intIndex1, _
             colScanned)               
            intIndex = intIndex1               
            Return(False)
        End If       
        If Not collectionCalculusWithReport_expressionParse__codeGen_ _
               (colParse, "0_PUSHLITERAL", intFrameSize) Then
            intIndex = intIndex1: Return(False)
        End If                                     
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' argumentListBody := arg [ , argumentListBody ]
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__argumentListBody_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer, _
                      ByRef intFrameSize As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "argumentListBody")                      
        Dim intIndex1 As Integer = intIndex 
        intFrameSize = 0  
        If Not collectionCalculusWithReport_expressionParse__arg_ _
               (colScanned, intIndex, colParse, intEndIndex) Then
            Return(False)
        End If     
        intFrameSize = 1                             
        If collectionCalculusWithReport_expressionParse__checkToken_ _
           (colScanned, intIndex, intEndIndex, "comma", ",") Then
            Dim intInnerFrameSize As Integer
            If Not collectionCalculusWithReport_expressionParse__argumentListBody_ _
                   (colScanned, intIndex, colParse, intEndIndex, intInnerFrameSize) Then
                collectionCalculusWithReport_expressionParse__errorHandler_ _
                ("Comma is not followed by an argument list continuation", _
                 "", _
                 intIndex, _
                 colScanned)
                intIndex = intIndex1: Return(False)                 
            End If                   
            intFrameSize += intInnerFrameSize
        End If                                  
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' byRefName := VB_IDENTIFIER_WITHUNDERSCORE
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__byRefName_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "byRefName")                      
        Dim intIndex1 As Integer = intIndex   
        Dim strName As String
        If Not collectionCalculusWithReport_expressionParse__checkToken_ _
               (strName, colScanned, intIndex, intEndIndex, "text") Then Return(False)
        If Mid(strName, 1, 1) <> "_" _
           OrElse _
           _OBJutilities.verify(strName, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_") <> 0 Then
            intIndex = intIndex1: Return(False)
        End If  
        If Not collectionCalculusWithReport_expressionParse__codeGen_ _
               (colParse, "1_PUSHBYREF", strName)       
            intIndex = intIndex1: Return(False)
        End If                             
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' calculusExpression := functionCall [ : calculusExpression ]
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__calculusExpression_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "calculusExpression")                      
        Dim intIndex1 As Integer = intIndex                      
        If Not collectionCalculusWithReport_expressionParse__functionCall_ _
               (colScanned, intIndex, colParse, intEndIndex) Then
            collectionCalculusWithReport_expressionParse__errorHandler_ _
            ("calculusExpression doesn't begin with a functionCall", _
             "BNF syntax expected: calculusExpression := functionCall [ : calculusExpression ]", _
             intIndex1, _
             colScanned)  
            Return(False)
        End If    
        If collectionCalculusWithReport_expressionParse__checkToken_ _
           (colScanned, intIndex, intEndIndex, "colon") Then
            If Not collectionCalculusWithReport_expressionParse__calculusExpression_ _
                   (colScanned, intIndex, colParse, intEndIndex) Then
                collectionCalculusWithReport_expressionParse__errorHandler_ _
                ("Colon isn't followed by a calculus expression or list of such expressions", _
                 "BNF syntax expected: calculusExpression := functionCall [ : calculusExpression ]", _
                 intIndex1, _
                 colScanned)  
                intIndex = intIndex1                          
                Return(False)
            End If    
        End If    
        Return(True)       
    End Function       
    
    ' ----------------------------------------------------------------------
    ' Check for specific tokens or types
    '
    '
    ' --- Check for a type: return its value
    Private Overloads Function collectionCalculusWithReport_expressionParse__checkToken_ _
                               (ByRef strTokenValue As String, _
                                ByVal colScanned As Collection, _
                                ByRef intIndex As Integer, _
                                ByVal intEndIndex As Integer, _
                                ByVal strTokenTypeExpected As String) As Boolean
        strTokenValue = ""                                
        If intIndex >= intEndIndex Then Return(False)                                
        Dim colEntry As Collection = CType(colScanned.Item(intIndex), Collection)                               
        If UCase(CStr(colEntry.Item(1))) = Trim(UCase(strTokenTypeExpected)) Then
            intIndex += 1
            strTokenValue = CStr(colEntry.Item(2))
            Return(True)
        End If          
        Return(False) 
    End Function    
    ' --- Check for a specific type: don't return its value                             
    Private Overloads Function collectionCalculusWithReport_expressionParse__checkToken_ _
                               (ByVal colScanned As Collection, _
                                ByRef intIndex As Integer, _
                                ByVal intEndIndex As Integer, _
                                ByVal strTokenTypeExpected As String) As Boolean
        Dim strValue As String                                
        Return(collectionCalculusWithReport_expressionParse__checkToken_ _
               (strValue, colScanned, intIndex, intEndIndex, strTokenTypeExpected))                                
    End Function    
    ' --- Check for a specific type and value                            
    Private Overloads Function collectionCalculusWithReport_expressionParse__checkToken_ _
                               (ByVal colScanned As Collection, _
                                ByRef intIndex As Integer, _
                                ByVal intEndIndex As Integer, _
                                ByVal strTokenTypeExpected As String, _
                                ByVal strTokenValueExpected As String) As Boolean
        Dim strValue As String 
        Dim intIndex1 As Integer = intIndex                               
        If Not collectionCalculusWithReport_expressionParse__checkToken_ _
               (strValue, colScanned, intIndex, intEndIndex, strTokenTypeExpected) Then
            Return(False)
        End If           
        If strValue <> strTokenValueExpected Then
            intIndex = intIndex1: Return(False)
        End If                                 
        Return(True)           
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Generate code for the calculus interpreter
    '
    '
    Private Overloads Function collectionCalculusWithReport_expressionParse__codeGen_ _
                               (ByRef colParse As Collection, _
                                ByVal strOp As String, _
                                ByVal objOperand As Object) As Boolean
        Dim colEntry As Collection
        Try
            colEntry = New Collection
            With colEntry
                .Add(strOp): .Add(objOperand)
            End With            
            colParse.Add(colEntry)
        Catch  
            collectionCalculusWithReport_errorHandler_("parse", _
                                                       "Cannot create entry collection or add it to code: " & _
                                                       Err.Number & " " & Err.Description, _
                                                       "Returning False")
            Return(False)                          
        End Try                                        
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' Append error information on behalf of the collectionCalculus methods
    '
    '
    Private Sub collectionCalculusWithReport_expressionParse__errorHandler_ _
                (ByVal strError As String, _
                 ByVal strHelp As String, _
                 ByVal intIndex As Integer, _
                 ByVal colScanned As Collection)  
        With colScanned
            Dim colEntry As Collection  
            Dim strPosition As String = "end of expression"
            Dim strToken As String  
            If intIndex <= .Count Then
                colEntry = CType(.Item(intIndex), Collection)  
                With colEntry
                    strPosition = "position " & CStr(.Item(3))
                    strToken = "of input expression " & _
                               "(" & CStr(.Item(1)) & " " & _
                               _OBJutilities.enquote(CStr(.Item(2))) & "): "
                End With         
            End If          
            collectionCalculusWithReport_errorHandler_("parse", _
                                                       "Error at " & _
                                                       strPosition & " " & _
                                                       strToken & ": " & _
                                                       strError, _
                                                       strHelp)               
        End With                         
    End Sub          
    
    ' ----------------------------------------------------------------------
    ' Locate the balancing parenthesis
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__findBalParenthesis_ _
                     (ByVal colScanned As Collection, _
                      ByVal intIndex As Integer, _
                      ByVal intEndIndex As Integer) As Integer
        Dim intLevel As Integer = 1
        Dim intIndex1 As Integer = intIndex
        Do While intIndex1 < intEndIndex
            If collectionCalculusWithReport_expressionParse__checkToken_ _
               (colScanned, intIndex1, intEndIndex, "parenthesis", "(") Then
                intLevel += 1
            ElseIf collectionCalculusWithReport_expressionParse__checkToken_ _
                   (colScanned, intIndex1, intEndIndex, "parenthesis", ")") Then
                intLevel -= 1
                If intLevel = 0 Then
                    Return(intIndex1 - 1)
                End If       
            Else
                intIndex1 += 1                         
            End If               
        Loop     
        collectionCalculusWithReport_expressionParse__errorHandler_ _
        ("Unbalanced parenthesis", "", intIndex1, colScanned)   
        Return(intIndex1)
    End Function                                                                     
    
    ' ----------------------------------------------------------------------
    ' methodName := COLLECTIONCALCULUS_METHOD (one of the supported
    '               method names) 
    '
    '
    ' Note: the list of supported method names in the following Select Case
    ' must be kept consistent both with the interpreter's list and with the
    ' documentation.
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__methodName_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer, _
                      ByRef strName As String) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "methodName")                      
        Dim intIndex1 As Integer = intIndex
        If Not collectionCalculusWithReport_expressionParse__checkToken_ _
               (strName, colScanned, intIndex, intEndIndex, "text") Then Return(False)                      
        Select Case UCase(strName)
            Case "COLLECTION2REPORT": 
            Case "COLLECTION2SET": 
            Case "COLLECTION2STRING": 
            Case "COLLECTION2STRINGREADABLE": 
            Case "COLLECTIONAPPEND": 
            Case "COLLECTIONCLEAR": 
            Case "COLLECTIONCOMPARE": 
            Case "COLLECTIONCOMPLEMENT": 
            Case "COLLECTIONCOPY": 
            Case "COLLECTIONDEPTH": 
            Case "COLLECTIONINTERSECTION": 
            Case "COLLECTIONITEMSET": 
            Case "COLLECTIONTYPES": 
            Case "COLLECTIONUNION": 
            Case "DEWEY2MEMBER": 
            Case "DUMPBYREF": 
            Case "EXTENDINDEX": 
            Case "ISINDEX": 
            Case "LIST2COLLECTION": 
            Case "MKRANDOMCOLLECTION": 
            Case "REM": 
            Case "SCALAR2COLLECTION": 
            Case "STRING2COLLECTION": 
            Case Else: 
                intIndex = intIndex1: Return(False)
        End Select  
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' functionCall := methodName [ argumentList ]
    '
    '
    Private Function collectionCalculusWithReport_expressionParse__functionCall_ _
                     (ByVal colScanned As Collection, _
                      ByRef intIndex As Integer, _
                      ByRef colParse As Collection, _
                      ByVal intEndIndex As Integer) As Boolean
        RaiseEvent collectionParseEvent(Me.Name, intIndex, "functionCall")                      
        Dim intIndex1 As Integer = intIndex                 
        Dim strName As String     
        If Not collectionCalculusWithReport_expressionParse__methodName_ _
               (colScanned, intIndex, colParse, intEndIndex, strName) Then
            Return(False)
        End If      
        Dim intIndex2 As Integer = intIndex
        If Not collectionCalculusWithReport_expressionParse__argumentList_ _
               (colScanned, intIndex, colParse, intEndIndex) Then
            If UCase(strName) = "REM" Then
                ' rem without parentheses should print the stack
                strName = "remStack"
            Else                
                collectionCalculusWithReport_expressionParse__codeGen_ _
                (colParse, "0_PUSHLITERAL", 0)
            End If                           
        End If               
        If Not collectionCalculusWithReport_expressionParse__codeGen_ _
               (colParse, strName, "") Then
            intIndex = intIndex1: Return(False)
        End If                     
        Return(True)
    End Function       
    
    ' ----------------------------------------------------------------------
    ' Scan expression on behalf of the collectionCalculus methods
    '
    '
    ' This method produces an unkeyed collection. Each item in this collection 
    ' is a subcollection with 2 items:
    '
    '
    '      *  Item(1) is the token type, one of the strings parenthesis,
    '         comma, colon, or text
    '
    '      *  Item(2) is the value of the token
    '
    '      *  Item(3) is its start index
    '
    '
    Private Function collectionCalculusWithReport_expressionScan_ _
                     (ByVal strExpression As String, _
                      ByRef strReport As String) _
            As Collection
        Dim colScanned As Collection
        Try
            colScanned = New Collection
        Catch  
            collectionCalculusWithReport_errorHandler_("scan", _
                                                       "Cannot create the scan collection: " & _
                                                       Err.Number & " " & Err.Description, _
                                                       "Returning Nothing")
            Return(Nothing)                          
        End Try   
        Try     
            With colScanned
                Dim intIndex1 As Integer = 1
                Dim intIndex2 As Integer
                Dim strCharacter As String
                Do While intIndex1 <= Len(strExpression)
                    ' --- Blank passover
                    For intIndex1 = intIndex1 To Len(strExpression)
                        If Mid(strExpression, intIndex1, 1) <> " " Then Exit For
                    Next intIndex1                
                    ' --- Locate comma, colon or parenthesis
                    intIndex2 = _OBJutilities.verify(strExpression & ")", ",:()", _
                                                    intIndex1, _
                                                    True)
                    ' --- Add open text if any
                    If intIndex2 > intIndex1 Then
                        If Not collectionCalculusWithReport_expressionScan__append_ _
                            (colScanned, _
                                "text", _
                                Mid(strExpression, intIndex1, intIndex2 - intIndex1), _
                                intIndex1) Then
                            Return(Nothing)
                        End If                           
                    End If       
                    ' --- May be finished
                    If intIndex2 > Len(strExpression) Then Exit Do  
                    intIndex1 = intIndex2      
                    ' --- Add comma, colon or parenthesis 
                    strCharacter = Mid(strExpression, intIndex1, 1) 
                    If Not collectionCalculusWithReport_expressionScan__append_ _
                            (colScanned, _
                            CStr(IIf(strCharacter = ",", _
                                    "comma", _
                                    CStr(IIf(strCharacter = ":", _
                                            "colon", _
                                            "parenthesis")))), _
                            strCharacter, _
                            intIndex1) Then
                        Return(Nothing)
                    End If      
                    intIndex1 += 1                     
                Loop    
                If Not (strReport Is Nothing) Then
                    _OBJutilities.append(strReport, _
                                        vbNewline, _
                                        "Scanned expression: " & _
                                        Me.collection2String(colScanned, _
                                                            booReadable:=True, _
                                                            booInspect:=False))       
                End If                                  
                Return(colScanned)
            End With        
        Catch
            collectionCalculusWithReport_errorHandler_("scan", _
                                                       Err.Number & " " & Err.Description, _
                                                       "")
        End Try            
    End Function                                                 
    
    ' ----------------------------------------------------------------------
    ' Append to collection, on behalf of expression scanner
    '
    '
    Private Function collectionCalculusWithReport_expressionScan__append_ _
                     (ByRef colScanned As Collection, _
                      ByVal strTokenType As String, _
                      ByVal strTokenValue As String, _
                      ByVal intTokenIndex As Integer) As Boolean
        Dim colEntry As Collection
        Try
            colEntry = New Collection
            With colEntry
                .Add(strTokenType)
                .Add(strTokenValue)
                .Add(intTokenIndex)
            End With            
            colScanned.Add(colEntry)            
        Catch  
            errorHandler_("Cannot create subcollection and/or append it to " & _
                          "the scan collection: " & _
                          Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_expressionScan__append_", _
                          "Destroying scan collection and returning False", _
                          True)
            Me.collectionClear(colScanned)
            Return(False)                                      
        End Try  
        Return(True)
    End Function 
    
    ' ----------------------------------------------------------------------
    ' Interpret the parsed collection calculus expression
    '
    '
    Private Function collectionCalculusWithReport_interpreter__ _
                     (ByVal strExpression As String, _
                      ByVal colParsed As Collection, _
                      ByVal colByRef As Collection, _
                      ByRef strReport As String) _
            As Object
        Dim objStack As Stack
        Try
            objStack = New Stack
        Catch ex As Exception
            errorHandler_("Cannot create stack: " & Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__", _
                          "Returning Nothing", _
                          True)
        End Try       
        Dim booByRef1 As Boolean
        Dim booByRef2 As Boolean
        Dim booResult As Boolean 
        Dim colHandle1 As Collection
        Dim colHandle2 As Collection
        Dim colNextEntry As Collection
        Dim intDepth As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intMax As Integer
        Dim intMin As Integer
        Dim intUBound As Integer
        Dim objHandle As Object
        Dim objPopValues() As Object
        Dim strName As String
        Dim strWork As String
        Dim strTypes As String
        If Not (strReport Is Nothing) Then
            strReport &= vbNewline & vbNewline & _
                         "Collection Calculus Evaluation at " & Now & vbNewline & vbNewline & _
                         "Expression: " & _
                         _OBJutilities.ellipsis(strExpression, 64)
        End If                         
        With colParsed
            Try
                For intIndex1 = 1 To .Count
                    colNextEntry = CType(.Item(intIndex1), Collection)
                    With colNextEntry
                        Select case Ucase(CStr(.Item(1)))
                            Case "0_PUSHLITERAL": 
                                objHandle = _OBJutilities.object2Scalar(.Item(2))
                                If (objHandle Is Nothing) Then objHandle = .Item(2)
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                    (objStack, objHandle) Then 
                                    Exit For
                                End If                                   
                            Case "1_PUSHBYREF": 
                                strName = CStr(.Item(2))
                                Try
                                    colHandle1 = CType(CType(colByRef(strName),Collection).Item(1), _
                                                       Collection)
                                Catch  
                                    errorHandler_("Undefined collection name " & strName, _
                                                  "collectionCalculusWithReport_interpreter__", _
                                                  "Evaluation will be ended")
                                    Exit For                                              
                                End Try                            
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                    (objStack, colHandle1) Then Exit For
                            Case "COLLECTION2REPORT": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    strName = Cstr(objPopValues(0))
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, strName)
                                    If (colHandle1 Is Nothing) Then Exit For                                        
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                   (objStack, Me.collection2Report(colHandle1)) Then 
                                    Exit For
                                End If                                    
                            Case "COLLECTION2SET": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For   
                                    booByRef1 = True
                                ElseIf (TypeOf objPopValues(0) Is Collection) Then
                                    colHandle1 = CType(objPopValues(0), Collection)
                                    booByRef1 = False
                                Else
                                    collectionCalculusWithReport_interpreter__errorArgs_
                                    Exit For                                    
                                End If           
                                If Not Me.collection2Set(colHandle1) Then Exit For
                                If Not booByRef1 _
                                   And _
                                   Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, colHandle1) Then
                                    Exit For
                                End If                                       
                            Case "COLLECTION2STRING":
                                If Not collectionCalculusWithReport_interpreter__collection2String_ _
                                       (objStack, colByRef, False) Then 
                                    Exit For
                                End If                                       
                            Case "COLLECTION2STRINGREADABLE":
                                If Not collectionCalculusWithReport_interpreter__collection2String_ _
                                       (objStack, colByRef, True) Then 
                                    Exit For
                                End If                                       
                            Case "COLLECTIONAPPEND": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                    (objStack, "collection|string collection|string", objPopValues) Then 
                                    Exit For
                                End If
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For   
                                    booByRef1 = True
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                    booByRef1 = False
                                End If           
                                If (TypeOf objPopValues(1) Is System.String) Then
                                    colHandle2 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(1)))
                                    If (colHandle2 Is Nothing) Then Exit For   
                                    booByRef2 = True
                                Else
                                    colHandle2 = CType(objPopValues(1), Collection)
                                    booByRef2 = False
                                End If           
                                If Not Me.collectionAppend(colHandle1, colHandle2) Then Exit For
                                If Not booByRef1 Then 
                                    collectionCalculusWithReport_interpreter__push_(objStack, colHandle1)
                                End If                                
                            Case "COLLECTIONCLEAR": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For   
                                    booByRef1 = True
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                    booByRef1 = False
                                End If           
                                If Not Me.collectionClear(colHandle1) Then Exit For
                                If Not booByRef1 _
                                   And _
                                   Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, colHandle1) Then
                                    Exit For
                                End If                                       
                            Case "COLLECTIONCOMPARE": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                    (objStack, "collection|string collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For                                        
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If (TypeOf objPopValues(1) Is System.String) Then
                                    colHandle2 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(1)))
                                    If (colHandle2 Is Nothing) Then Exit For                                        
                                Else
                                    colHandle2 = CType(objPopValues(1), Collection)
                                End If           
                                booResult = Me.collectionCompare(colHandle1, colHandle2)                                                 
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, booResult) Then Exit For
                            Case "COLLECTIONCOMPLEMENT": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If (TypeOf objPopValues(1) Is System.String) Then
                                    colHandle2 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(1)))
                                    If (colHandle2 Is Nothing) Then Exit For    
                                Else
                                    colHandle2 = CType(objPopValues(1), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                        (objStack, Me.collectionComplement(colHandle1, colHandle2)) Then 
                                    Exit For
                                End If   
                            Case "COLLECTIONCOPY": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string [ string ]", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                colHandle2 = Me.collectionCopy(colHandle1)
                                If UBound(objPopValues) = 1 Then
                                    If Not collectionCalculusWithReport_interpreter__saveByRef_ _
                                           (colByRef, CStr(objPopValues(1)), colHandle2) Then
                                        Exit For
                                    End If      
                                Else                                                                         
                                    If Not collectionCalculusWithReport_interpreter__push_ _
                                           (objStack, colHandle2) Then 
                                        Exit For
                                    End If   
                                End If                                                                     
                            Case "COLLECTIONDEPTH": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For                                        
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                    (objStack, Me.collectionDepth(colHandle1)) Then Exit For
                            Case "COLLECTIONFIND": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string object", objPopValues) Then 
                                    Exit For
                                End If
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For                                        
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.collectionFind(colHandle1, objPopValues(1))) Then 
                                    Exit For
                                End If                                    
                            Case "COLLECTIONINTERSECTION": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If (TypeOf objPopValues(1) Is System.String) Then
                                    colHandle2 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(1)))
                                    If (colHandle2 Is Nothing) Then Exit For    
                                Else
                                    colHandle2 = CType(objPopValues(1), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.collectionIntersection(colHandle1, colHandle2)) Then 
                                    Exit For
                                End If   
                            Case "COLLECTIONITEMSET": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string string|integer object", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.collectionItemSet(colHandle1, _
                                                                       _OBJutilities.dequote(CStr(objPopValues(1))), _
                                                                       objPopValues(2))) Then 
                                    Exit For
                                End If   
                            Case "COLLECTIONSAVE": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__saveByRef_ _
                                       (colByRef, CStr(objPopValues(1)), colHandle1) Then
                                    Exit For
                                End If      
                            Case "COLLECTIONTYPES": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string [ string ]", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                colHandle2 = Me.collectionTypes(colHandle1)
                                If UBound(objPopValues) = 1 Then
                                    strName = CStr(objPopValues(1))
                                    If Not collectionCalculusWithReport_interpreter__saveByRef_ _
                                           (colByRef, strName, colHandle2) Then
                                        Exit For
                                    End If      
                                Else                                                                         
                                    If Not collectionCalculusWithReport_interpreter__push_ _
                                           (objStack, colHandle2) Then 
                                        Exit For
                                    End If   
                                End If                                                                     
                            Case "COLLECTIONUNION": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string collection|string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If (TypeOf objPopValues(1) Is System.String) Then
                                    colHandle2 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(1)))
                                    If (colHandle2 Is Nothing) Then Exit For    
                                Else
                                    colHandle2 = CType(objPopValues(1), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.collectionUnion(colHandle1, colHandle2)) Then 
                                    Exit For
                                End If   
                            Case "DEWEY2MEMBER": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection|string string", objPopValues) Then Exit For
                                If (TypeOf objPopValues(0) Is System.String) Then
                                    colHandle1 = _
                                        collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                                        (colByRef, Cstr(objPopValues(0)))
                                    If (colHandle1 Is Nothing) Then Exit For    
                                Else
                                    colHandle1 = CType(objPopValues(0), Collection)
                                End If           
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.dewey2Member(colHandle1, objPopValues(1))) Then 
                                    Exit For
                                End If   
                            Case "DUMPBYREF":
                                _OBJutilities.append(strReport, _
                                                     vbNewline, _
                                                     _OBJutilities.string2Box _
                                                     (_OBJutilities.soft2HardParagraph _
                                                      (Me.collection2String(colByRef), _
                                                       bytLineWidth:=50), _
                                                      "BY REFERENCE COLLECTION")) 
                            Case "EXTENDINDEX":                                                      
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection string object", objPopValues) Then Exit For
                                If Not Me.extendIndex(CType(objPopValues(0), Collection), _
                                                      objPopValues(1), _
                                                      objPopValues(2)) Then
                                    Exit For
                                End If                                                                                             
                            Case "ISINDEX":
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "collection", objPopValues) Then Exit For
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, Me.isIndex(CType(objPopValues(0), Collection))) Then             
                                    Exit For
                                End If                                                                                                                                                                                   
                            Case "LIST2COLLECTION": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "object*", objPopValues) Then 
                                    Exit For
                                End If   
                                intUBound = -1
                                Try
                                    intUBound = UBound(objPopValues)
                                Catch: End Try                                
                                Select Case UBound(objPopValues)
                                    Case -1:
                                        colHandle1 = Me.list2Collection()
                                    Case 0:
                                        colHandle1 = Me.list2Collection(objPopValues(0))
                                    Case 1:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1))
                                    Case 2:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2))
                                    Case 3:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3))
                                    Case 4:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4))
                                    Case 5:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4), _
                                                                        objPopValues(5))
                                    Case 6:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4), _
                                                                        objPopValues(5), _
                                                                        objPopValues(6))
                                    Case 7:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4), _
                                                                        objPopValues(5), _
                                                                        objPopValues(6), _
                                                                        objPopValues(7))
                                    Case 8:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4), _
                                                                        objPopValues(5), _
                                                                        objPopValues(6), _
                                                                        objPopValues(7), _
                                                                        objPopValues(8))
                                    Case 9:
                                        colHandle1 = Me.list2Collection(objPopValues(0), _
                                                                        objPopValues(1), _
                                                                        objPopValues(2), _
                                                                        objPopValues(3), _
                                                                        objPopValues(4), _
                                                                        objPopValues(5), _
                                                                        objPopValues(6), _
                                                                        objPopValues(7), _
                                                                        objPopValues(8), _
                                                                        objPopValues(9))
                                    Case Else:
                                        errorHandler_("Too many list elements in list2Collection", _
                                                      "collectionCalculusWithReport_interpreter__", _
                                                      "Creating empty collection")   
                                        colHandle1 = Me.list2Collection()                                                                                                                           
                                End Select                                
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, colHandle1) Then 
                                    Exit For
                                End If                                    
                            Case "MKRANDOMCOLLECTION": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "integer integer string integer [ string ]", objPopValues) Then
                                    Exit For
                                End If   
                                colHandle1 = Me.mkRandomCollection(CInt(objPopValues(0)), _
                                                                   CInt(objPopValues(1)), _
                                                                   _OBJutilities.dequote(CStr(objPopValues(2))), _
                                                                   CInt(objPopValues(3)))
                                If UBound(objPopValues) = 4 Then
                                    If Not collectionCalculusWithReport_interpreter__saveByRef_ _
                                           (colByRef, CStr(objPopValues(4)), colHandle1) Then 
                                        Exit For
                                    End If
                                Else                                                                               
                                    If Not collectionCalculusWithReport_interpreter__push_ _
                                            (objStack, colHandle1) Then 
                                        Exit For
                                    End If  
                                End If                                                                                                                                     
                            Case "REM": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                        (objStack, "string", objPopValues) Then 
                                    Exit For
                                End If  
                                If Not (strReport Is Nothing) Then
                                    _OBJutilities.append(strReport, vbNewline, CStr(objPopValues(0)))
                                End If                                    
                            Case "REMSTACK": 
                                If Not (strReport Is Nothing) Then
                                    strWork = "<Empty stack>"
                                    With objStack
                                        If .Count > 0 Then strWork = object2Display_(.Peek)
                                    End With                     
                                    _OBJutilities.append(strReport, vbNewline, strWork)
                                End If                                    
                            Case "STRING2COLLECTION": 
                                If Not collectionCalculusWithReport_interpreter__popFrame_ _
                                       (objStack, "string", objPopValues) Then 
                                    Exit For
                                End If   
                                colHandle1 = Me.string2Collection(strWork)
                                If Not collectionCalculusWithReport_interpreter__push_ _
                                       (objStack, colHandle1) Then Exit For
                        End Select      
                        objHandle = Nothing
                        With objStack
                            If .Count > 0 Then objHandle = .Peek
                        End With                    
                        RaiseEvent collectionCalculusEvent(Me.Name, CStr(.Item(1)), objHandle)              
                    End With                
                Next intIndex1            
                If intIndex1 <= .Count OrElse objStack.Count < 1 Then 
                    Return(Nothing)
                End If            
                Return(objStack.Peek)    
            Catch
                collectionCalculusWithReport_errorHandler_("interpreter", _
                                                            "Error in interpreter: " & _
                                                            Err.Number & " " & Err.Description & vbNewline & _
                                                            "Occurs at op " & intIndex1 & _
                                                            ": " & _
                                                            Me.collection2String _
                                                            (CType(colParsed.Item(intIndex1), Collection), _
                                                             booReadable:=True, booInspect:=False), _
                                                            "Returning Nothing")
                Return(Nothing)                                      
            End Try
        End With        
    End Function   
    
    ' ----------------------------------------------------------------------
    ' Interpreter's collection2String interface
    '
    '
    Private Function collectionCalculusWithReport_interpreter__collection2String_ _
                     (ByRef objStack As Stack, _
                      ByVal colByRef As Collection, _
                      ByVal booReadable As Boolean) _
            As Boolean
        Dim colHandle As Collection            
        Dim objPopValues() As Object                        
        If Not collectionCalculusWithReport_interpreter__popFrame_ _
               (objStack, "collection|string", objPopValues) Then Return(False)
        If (TypeOf objPopValues(0) Is System.String) Then
            colHandle = _
                collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                (colByRef, Cstr(objPopValues(0)))
            If (colHandle Is Nothing) Then Return(False)                                       
        Else
            colHandle = CType(objPopValues(0), Collection)
        End If           
        If Not collectionCalculusWithReport_interpreter__push_ _
               (objStack, _
                Me.collection2String(colHandle, _
                                     booReadable:=booReadable, _
                                     booInspect:=Not booReadable)) Then 
            Return(False)
        End If            
        Return(True)                                    
    End Function            
    
    ' ----------------------------------------------------------------------
    ' Dereference collection identified by its byref name
    '
    '
    Private Function collectionCalculusWithReport_interpreter__dereferenceCollection_ _
                     (ByVal colByRef As Collection, _
                      ByVal strName As String) As Collection
        Try
            Return(CType(CType(colByRef.Item(strName), Collection).Item(1), _
                         Collection))
        Catch  
            collectionCalculusWithReport_errorHandler_("interpreter", _
                                                       "Collection name " & _
                                                       _OBJutilities.enquote(strName) & " " & _
                                                       "not found", _
                                                       "Returning Nothing")
            Return(Nothing)
        End Try        
    End Function     
    
    ' ----------------------------------------------------------------------
    ' Report an argument error
    '
    '
    Private Function collectionCalculusWithReport_interpreter__errorArgs_ _
            As Boolean
        errorHandler_("Unexpected argument syntax", _
                      "collectionCalculusWithReport_interpreter_", _
                      "Terminating evaluation")
        Return(False)                      
    End Function            
    
    ' ----------------------------------------------------------------------
    ' Pop 0..n operands preceded by arglist count on behalf of the interpreter
    '
    '
    ' Note that strFrameExpected specifies the expected frame as a series of
    ' blank-delimited words.
    '
    ' Each element of strFrameExpected may be one type such as string or
    ' collection, or a set of two or more alternative types, separated by
    ' a vertical bar.
    '
    ' Also, the optional part of the expected frame, if any, should be in
    ' square brackets, and at the end of the expected frame.
    '
    ' If the last element may repeat indefinitely it should be followed by
    ' a space and then an asterisk.
    '
    '
    Private Overloads Function collectionCalculusWithReport_interpreter__popFrame_ _
                               (ByRef objStack As Stack, _
                                ByVal strFrameExpected As String, _
                                ByRef objPopValues() As Object) As Boolean
        Dim intMax As Integer                                
        Dim intMin As Integer                                
        Dim colFrameExpected As Collection = _                                
            collectionCalculusWithReport_interpreter__popFrame__frame2Collection_ _
            (strFrameExpected, intMin, intMax)
        If (colFrameExpected Is Nothing) Then Return(False)
        Dim intCount As Integer          
        Try
            intCount = CInt(objStack.Pop)
        Catch  
            errorHandler_("Internal error, stack frame format is not valid or " & _
                          "error in stack pop: " & _
                          Err.Number & " " & Err.Description, _
                          "", _
                          "Marking object not usable and returning False", _
                          True)
            Return(False)                           
        End Try           
        If intCount < intMin OrElse intMax >= 0 AndAlso intCount > intMax Then
            Return(collectionCalculusWithReport_interpreter__errorArgs_)
        End If        
        Try
            Redim objPopValues(intCount - 1)
        Catch ex As Exception
            errorHandler_("Cannot allocate pop value array: " & _
                          Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__popFrame_", _
                          "Returning False", _
                          True)
            Return(False)                          
        End Try    
        Dim colEntry As Collection
        Dim intIndex1 As Integer
        Dim strNext As String
        For intIndex1 = intCount - 1 To 0 Step -1
            If intIndex1 + 1 > objStack.Count Then
                Return(collectionCalculusWithReport_interpreter__errorArgs_)
            End If            
            Try
                objPopValues(intIndex1) = objStack.Pop
            Catch  
                errorHandler_("Cannot pop: " & Err.Number & " " & Err.Description, _
                              "collectionCalculusWithReport_interpreter__pop_", _
                              "Returning False: the pop value is Nothing", _
                              True)
                Return(False)                          
            End Try 
        Next intIndex1              
        Dim intIndex2 As Integer
        Dim intWork As Integer
        Dim strWork As String
        For intIndex1 = 0 To UBound(objPopValues)
            colEntry = CType(colFrameExpected.Item(Math.Min(intIndex1 + 1, colFrameExpected.Count)), _
                             Collection) 
            With colEntry
                For intIndex2 = 1 To .Count
                    Select Case UCase(CStr(.Item(intIndex2)))
                        Case "COLLECTION": 
                            If (TypeOf objPopValues(intIndex1) Is Collection) Then 
                                Exit For
                            End If
                        Case "STRING": 
                            Try
                                strWork = CStr(objPopValues(intIndex1))
                                Exit For
                            Catch: End Try                            
                        Case "INTEGER": 
                            Try
                                intWork = CInt(objPopValues(intIndex1))
                                Exit For
                            Catch: End Try                            
                        Case "OBJECT": 
                            Exit For
                        Case Else: 
                            errorHandler_("Internal programming error: " & _
                                          "Unsupported type " & _
                                          _OBJutilities.enquote(CStr(.Item(intIndex2))), _
                                          "collectionCalculusWithReport_interpreter__popFrame_", _
                                          "Returning False")
                            Return(False)                                          
                    End Select             
                Next intIndex2  
                If intIndex2 > .Count Then
                    Return(collectionCalculusWithReport_interpreter__errorArgs_)                                          
                End If                              
            End With                                   
            intIndex2 -= 1
        Next intIndex1          
        Return(True)
    End Function                          
    
    ' ----------------------------------------------------------------------
    ' Parse frame expression
    '
    '
    Private Function collectionCalculusWithReport_interpreter__popFrame__frame2Collection_ _
                     (ByVal strFrameExpression As String, _
                      ByRef intMin As Integer, _
                      ByRef intMax As Integer) _
            As Collection
        Dim booHasBrackets As Boolean            
        Dim intIndex1 As Integer        
        Dim intIndex2 As Integer        
        Dim strFrameExpressionWork As String = Trim(strFrameExpression)   
        Dim strSplit1() As String 
        Dim strSplit2() As String 
        intIndex1 = Instr(strFrameExpressionWork & "[", "[")
        booHasBrackets = (intIndex1 <= Len(strFrameExpressionWork))
        intIndex2 = Len(strFrameExpressionWork) + 1
        If intIndex1 <= Len(strFrameExpressionWork) Then
            If Mid(strFrameExpressionWork, Len(strFrameExpressionWork)) <> "]" Then
                errorHandler_("Internal programming error in a frame expression: " & _
                              "unbalanced bracket", _
                              "collectionCalculusWithReport_interpreter__popFrame__frame2Collection_", _
                              "Returning Nothing")
                Return(Nothing)                              
            End If            
            intIndex2 = Len(strFrameExpressionWork)
        End If  
        Try              
            strSplit1 = split(Trim(Mid(strFrameExpressionWork, 1, intIndex1 - 1)), " ")
            strSplit2 = split(Trim(Mid(strFrameExpressionWork, _
                                       intIndex1 + 1, _
                                       Math.Max(intIndex2 - intIndex1 - 1, 0))), _
                              " ")
        Catch
            errorHandler_("Error in split: " & Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__popFrame__frame2Collection_", _
                          "Returning Nothing")
            Return(Nothing)                              
        End Try
        intMin = -1
        Dim booHasAsterisk As Boolean
        Dim booOK As Boolean
        Dim colNextEntry As Collection
        Dim colNew As Collection = Me.list2Collection()
        If (colNew Is Nothing) Then Return(Nothing)
        For intIndex1 = 0 To UBound(strSplit1)
            colNextEntry = _
            collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_ _
            (strSplit1(intIndex1), booHasAsterisk)   
            If (colNextEntry Is Nothing) Then Exit For
            If booHasAsterisk Then
                If intIndex1 < UBound(strSplit1) OrElse booHasBrackets Then
                    errorHandler_("Internal, programming error: misplaced asterisk in frame expression", _
                                  "collectionCalculusWithReport_interpreter__popFrame__frame2Collection_", _
                                  "Asterisk will be ignored")
                Else
                    intMax = -1  
                    If intIndex1 = 0 Then intMin = 0                                
                End If                                  
            End If     
            If colNextEntry.Count > 0 Then
                colNew.Add(colNextEntry)
            End If                                 
        Next intIndex1       
        If intMax <> -1 Then intMax = intMin
        If intIndex1 > UBound(strSplit1) Then
            For intIndex1 = 0 To UBound(strSplit2)
                colNextEntry = _
                collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_ _
                (strSplit2(intIndex1), booHasAsterisk)   
                If (colNextEntry Is Nothing) Then Exit For
                If booHasAsterisk Then
                    If intIndex1 < Ubound(strSplit2) Then
                        errorHandler_("Internal programming error: asterisk is misplaced " & _
                                      "in frame syntax", _
                                      "collectionCalculusWithReport_interpreter__popFrame__frame2Collection_", _
                                      "Asterisk will be ignored")
                    Else                                      
                        intMax = -1                             
                    End If                    
                End If                          
                If colNextEntry.Count > 0 Then
                    colNew.Add(colNextEntry)
                    If intMax <> -1 Then intMax += 1
                End If                                 
            Next intIndex1       
            booOK = (intIndex2 > UBound(strSplit2))
        End If         
        If Not booOK Then
            Me.collectionClear(colNew): Return(Nothing)
        End If        
        If intMin = -1 Then intMin = UBound(strSplit1) + 1
        Return(colNew)
    End Function          
    
    ' ----------------------------------------------------------------------
    ' Create frame collection entry as collection of alternatives 
    '
    '
    Private Function collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_ _
                     (ByVal strAlternatives As String, _
                      ByRef booHasAsterisk As Boolean) _
                     As Collection
        Dim strSplit() As String
        Try
            strSplit = split(strAlternatives, "|")
        Catch  
            errorHandler_("Can't split: " & Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_", _
                          "Returning Nothing")
            Return(Nothing)                          
        End Try     
        Dim colEntry As Collection = Me.list2Collection()
        If colEntry Is Nothing Then Return(Nothing)
        Dim intIndex1 As Integer
        Dim strNext As String
        booHasAsterisk = False
        Try
            With colEntry
                For intIndex1 = 0 To UBound(strSplit)
                    strNext = strSplit(intIndex1)
                    If strNext <> "" Then
                        If Mid(strNext, Len(strNext)) = "*" Then
                            If Len(strNext) < 2 OrElse intIndex1 < UBound(strSplit) Then
                                errorHandler_("Internal, programming error in frame: misplaced asterisk", _
                                            "collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_", _
                                            "Returning Nothing")
                                Return(Nothing)                                          
                            End If       
                            booHasAsterisk = True
                            strNext = Mid(strNext, 1, Len(strNext) - 1)
                        End If                    
                        .Add(strNext)
                    End If
                Next intIndex1            
            End With       
        Catch
            errorHandler_("Can't create collection entry: " & Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__popFrame__frame2Collection__mkEntry_", _
                          "Returning Nothing")
            Return(Nothing)                          
        End Try 
        Return(colEntry)               
    End Function                                                
    
    ' ----------------------------------------------------------------------
    ' Push on behalf of the interpreter
    '
    '
    Private Function collectionCalculusWithReport_interpreter__push_ _
                     (ByRef objStack As Stack, ByVal objObject As Object) As Boolean
        Try
            objStack.Push(objObject)
        Catch  
            errorHandler_("Cannot push: " & Err.Number & " " & Err.Description, _
                          "collectionCalculusWithReport_interpreter__push_", _
                          "Returning False", _
                          True)
            Return(False)                          
        End Try        
        Return(True)
    End Function              
    
    ' ----------------------------------------------------------------------
    ' Save Byref, named collection
    '
    '
    Private Function collectionCalculusWithReport_interpreter__saveByRef_ _
                     (ByRef colByRef As Collection, _
                      ByVal strName As String, _
                      ByRef colCollection As Collection) _
                     As Boolean
        With colByRef        
            Dim colHandle As Collection
            Dim booExists As Boolean = True             
            Try
                colHandle = CType(CType(.Item(strName), Collection).Item(1), Collection) 
            Catch  
                booExists = False
            End Try        
            If booExists Then
                .Remove(strName)
                collectionClear(colHandle)
            End If                
            Try
                colHandle = New Collection
                With colHandle
                    .Add(colCollection): .Add(strName)
                End With                
                .Add(colHandle, strName) 
            Catch  
                errorHandler_("Cannot replace named collection " & strName & ": " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              "Returning False", _
                              True)
                Return(False)                              
            End Try        
            Return(True)
        End With
    End Function                                                          

#End Region ' collectionCalculus

End Class
