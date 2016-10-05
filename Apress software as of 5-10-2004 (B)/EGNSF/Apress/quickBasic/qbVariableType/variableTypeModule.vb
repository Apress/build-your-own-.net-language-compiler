Option Strict

Namespace variableTypes
    Module variableTypeModule
        ' ***** Variable types as constant list *****
        Public Const VARIABLETYPES As String = _
        "Boolean Byte Integer Long Single Double String Variant Array Null Unknown"
        ' ***** Variable types as enumerator *****
        Public Enum ENUvarType
            vtBoolean           ' Boolean
            vtByte              ' 8 bit unsigned integer
            vtInteger           ' 16 bit signed integer
            vtLong              ' 32 bit signed integer (VB.Net Long not supported)
            vtSingle            ' Single precision (we use VB.Net double precision)
            vtDouble            ' Double precision (uses VB.Net doubles, does not accurately
                                ' reflect QuickBasic precision)
            vtString            ' Variable length string
            vtVariant           ' Object 
            vtArray             ' Array
            vtUnknown           ' Default 
            vtNull              ' Null: used when the VarType is requested of a noncontainer 
        End Enum
    End Module
End NameSpace
