Attribute VB_Name = "mStableQuick"
Option Explicit                           ' -©Rd 2006/2008-

' + Stable QuickSort 2.3 +++++++++++++++++++++++++++++++++++++++++++

' These sorting routines have the following features:

' - They can handle sorting arrays of millions of string items.
' - They can handle sorting in ascending and descending order.
' - They can handle case-sensitive and case-insensitive criteria.
' - They can handle zero or higher based arrays.
' - They can handle negative lb and positive ub.
' - They can handle negative lb and zero or negative ub.
' - They can sort sub-sets of the array data.

' + Background +++++++++++++++++++++++++++++++++++++++++++++++++++++

' This is a non-recursive quicksort based algorithm that has been written from
' the ground up as a stable alternative to the blindingly fast quicksort.

' It is not quite as fast as the outright fastest non-stable quicksort, but is
' still very fast as it uses buffers and copymemory and is beaten by none of my
' other string sorting algorithms except my fastest non-stable quicksort.

' A standard quicksort only moves items that need swapping, while this stable
' algorithm manipulates all items on every iteration to keep them all in relative
' positions to one another. This algorithm I have dubbed the Avalanche©.

' It is also important to note that this algorithm does not suffer at all from
' the traditional quicksort nemesis. It is in fact much faster at re-sorting data
' that has been pre-sorted, and sorting data with many repeated items, than most
' other sorting algorithms as it has been highly optimized for these data states!

' + Version 2.1 ++++++++++++++++++++++++++++++++++++++++++++++++++++

' This is a re-working of my stable quicksort algorithm.

' A runner section has been added to handle a very hard job for a stable sorter;
' reverse pretty sorting. This is case-insensitive sorting of data that has been
' pre-sorted case-sensitively in reverse order - lower-case first in ascending
' order, or capitals first in descending order.

' It utilises a runner technique to boost this very demanding operation,
' down from 2.0 to 1.5 seconds on 100,000 items on my 866 MHz P3.
' Adding runners has also boosted same-direction pretty sorting operations.

' Because all items are re-positioned based on the current value it can identify
' when the avalanche process is producing a zero count buffer one way and so is
' moving all items the other way, indicating that the data is in a pre-sorted state
' (shifting no items up/down in relation to the current item).

' On each iteration a test of the buffer counts can identify when it is re-sorting
' or reverse-sorting, as well as producing distinctive indicators on reverse-pretty
' and same-direction pretty sorting operations. The range becomes very small on
' unsorted data before a small range produces a zero count buffer, so small ranges
' are ignored to skip false indicators.

' So when performing a pretty or reverse-pretty operation the code can identify
' this state and the runners are turned on automatically.

' + Version 2.2 ++++++++++++++++++++++++++++++++++++++++++++++++++++

' This version identifies sub-sets of pre-sorted data and delegates it to
' a built-in insert/binary hybrid algorithm dubbed the Twister©.

' This delegation is the sole reason for the speed boost on all operations
' over version 2.1, and also the reason for the incredibly fast refresh
' sorting performance - it can refresh-sort 3,248,230 pre-sorted strings
' in around 2 and a half seconds on my 866MHz P3.

' + Version 2.25 +++++++++++++++++++++++++++++++++++++++++++++++++++

' Interim version 2.25 added safe addition and subtraction of unsigned
' long integers.

' This guarantees safe arithmetic operations on memory address pointers
' which are used extensively by the runner sections of code.

' This change imposed a slight performance degradation on all operations.

' + Version 2.3 ++++++++++++++++++++++++++++++++++++++++++++++++++++

' The latest version of this algorithm employs a SAFEARRAY substitution
' technique to trick VB into thinking the four-byte string pointers in
' the string array are just VB longs in a native VB long array.

' The technique simply uses CopyMemory to point a VB long array (defined
' in the module) at the first of the string pointers in memory, and sets
' its lower-bound and item count to match (as if it had been redimmed).

' This allows us to treat the string pointers as if they were simply
' four-byte long values in a long array and can be swapped around as
' needed without touching the actual strings that are pointed to.

' Reading and assigning to a VB long array is lightning fast, and proves
' to be considerably faster when copying only one item than the previous
' method of copying the string pointers using CopyMemory.

' This stable algorithm is truely very fast at all sorting operations!

' + Indexed Version ++++++++++++++++++++++++++++++++++++++++++++++++

' This version receives a dynamic long array that holds references to the
' string arrays indices. This is known as an indexed sort. No changes are
' made to the source string array.

' The index array is automatically initialized if it is passed erased or
' uninitialized. The index array can be passed again for sorting without
' erasing it.

' After a sort procedure is run the long array is ready as a sorted index
' (lookup table) to the string array items.

' E.G strA(idxA(lo)) returns the lo item in the string array whose index
' may be anywhere in the string array.

' Usage Details:

' The index array can be redimmed to match the source string array boundaries
' or it can be erased or left uninitialized before sorting a string array for
' the first time. However, if you modify string items and re-sort you should
' not redim or erase the index array which will take advantage of the fast
' refresh sorting performance. This also allows the index array to be passed
' on to other sorting processes to be further manipulated.

' Even when using redim with the preserve keyword and adding more items to the
' string array you can pass the index array unchanged and the new items will be
' sorted into the previously sorted array. The index array will automatically
' return with boundaries matching the string array boundaries.

' Only when you reload the string array items with new array boundaries should
' you erase the index array for the first sorting operation. Also, if you redim
' the source string array to smaller boundaries you should erase the index array
' before sorting the new smaller data set for the first time.

' See the header comments for ValidateIndexArray for more details.

' + Licence Agreement ++++++++++++++++++++++++++++++++++++++++++++++

' You are free to use any part or all of this code even for commercial
' purposes in any way you wish under the one condition that no copyright
' notice is moved or removed from where it is.

' For comments, suggestions or bug reports you can contact me at:
' rd•edwards•bigpond•com.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Declare some CopyMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

' More efficient repeated use of numeric literals
Private Const n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&, n6 = 6&
Private Const n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&, n64 = 64&
Private Const n10K As Long = 10000&
Private Const n20K As Long = 20000&
Private Const n50K As Long = 50000

Private Const rRunner4 As Single = 0.0025 '0.002<<reverse-sorting-0.003-unsorted>>0.004
Private Const rRunner5 As Single = 0.0015 '0.001<<reverse-sorting-unsorted>>0.002

' Used for unsigned arithmetic
Private Const DW_MSB = &H80000000 ' DWord Most Significant Bit

Private Enum SAFEATURES
    FADF_AUTO = &H1               ' Array is allocated on the stack
    FADF_STATIC = &H2             ' Array is statically allocated
    FADF_EMBEDDED = &H4           ' Array is embedded in a structure
    FADF_FIXEDSIZE = &H10         ' Array may not be resized or reallocated
    FADF_BSTR = &H100             ' An array of BSTRs
    FADF_UNKNOWN = &H200          ' An array of IUnknown*
    FADF_DISPATCH = &H400         ' An array of IDispatch*
    FADF_VARIANT = &H800          ' An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8    ' Bits reserved for future use
    #If False Then
        Dim FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
    #End If
End Enum
Private Const VT_BYREF = &H4000&  ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix
Private Const FADF_NO_REDIM = FADF_AUTO Or FADF_FIXEDSIZE

Private Type SAFEARRAY
    cDims       As Integer        ' Count of dimensions in this array
    fFeatures   As Integer        ' Bitfield flags indicating attributes of a particular array
    cbElements  As Long           ' Byte size of each element of the array
    cLocks      As Long           ' Number of times the array has been locked without corresponding unlock
    pvData      As Long           ' Pointer to the start of the array data (use only if cLocks > 0)
    cElements   As Long           ' Count of elements in this dimension
    lLBound     As Long           ' The lower-bounding index of this dimension
    lUBound     As Long           ' The upper-bounding index of this dimension
End Type

Private StringPtrs_Header As SAFEARRAY
Private StringPtrs() As Long

Private ssLb() As Long, ssUb() As Long, ssMax As Long  ' Avalanche pending boundary stacks
Private psLb() As Long, psUb() As Long, psMax As Long  ' Stable presorter boundary stacks
Private lA_1() As Long, lA_2() As Long, ssBuf As Long  ' Stable quicksort working buffers
Private twLb() As Long, twUb() As Long, twMax As Long  ' Twister runner stacks
Private twBuf() As Long, twBufMax As Long              ' Twister copymemory buffer

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Enum eSortOrder
    Descending = -1&
    Default = 0&
    Ascending = 1&
    #If False Then
        Dim Descending, Default, Ascending
    #End If
End Enum

Public Enum eCompareMethod
    BinaryCompare = &H0
    TextCompare = &H1
    #If False Then
        Dim BinaryCompare, TextCompare
    #End If
End Enum

Public Enum eCompareResult
    Lesser = -1&
    Equal = 0&
    Greater = 1&
    #If False Then
        Dim Lesser, Equal, Greater
    #End If
End Enum

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const Default_Order As Long = Ascending
Private mMethod As eCompareMethod
Private mOrder As eSortOrder

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The following properties should be set before sorting.

Property Get SortOrder() As eSortOrder
    If mOrder = Default Then mOrder = Default_Order
    SortOrder = mOrder
End Property

Property Let SortOrder(ByVal eNewOrder As eSortOrder)
    If eNewOrder = Default Then
        If mOrder = Default Then mOrder = Default_Order
    Else
        mOrder = eNewOrder
    End If
End Property

Property Get SortMethod() As eCompareMethod
    SortMethod = mMethod
End Property

Property Let SortMethod(ByVal eNewMethod As eCompareMethod)
    mMethod = eNewMethod
End Property

' + Pretty Sort ++++++++++++++++++++++++++++++++++++++++++++

' Sort with binary comparison to seperate upper and lower
' case letters in the order specified by CapsFirst.

' Then sort in the desired direction with case-insensitive
' comparison to group upper and lower case letters together,
' but with a stable sort to preserve the requested caps-first
' or lower-first order.

Sub strPrettySort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True)
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim eMethod As eCompareMethod
    Dim eOrder As eCompareResult

    eMethod = SortMethod
    eOrder = SortOrder

    If CapsFirst Then SortOrder = Ascending Else SortOrder = Descending
    SortMethod = BinaryCompare

    strStableSort2 sA, lbA, ubA

    SortOrder = eOrder    ' Reset SortOrder property
    SortMethod = TextCompare

    strStableSort2 sA, lbA, ubA

    SortMethod = eMethod  ' Reset SortMethod property
End Sub

' + Stable QuickSort v2.3 +++++++++++++++++++++++++++++++++++++++++++++

' This is a non-recursive quicksort based algorithm that has been written from
' the ground up as a stable alternative to the blindingly fast quicksort.

Sub strStableSort2(sA() As String, ByVal lbA As Long, ByVal ubA As Long)
    ' This is an even faster stable non-recursive quicksort
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim item As String, lpStr As Long, lpS As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim idx As Long, opt As Long, pvt As Long
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, run As Long, cast As Long
    Dim ceil As Long, mezz As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim lItem As Long, lpSA As Long
    Dim eComp As eSortOrder

    cnt = ubA - lbA + n1              ' Grab array item count
    If (cnt < n2) Then Exit Sub       ' If nothing to do then exit
    eComp = SortOrder                 ' Initialize compare variable
    pvt = (cnt \ n64) + n32           ' Allow for worst case senario + some

    lpSA = SubstArrayHeader(sA, lbA, ubA)
    InitializeStacks ssLb, ssUb, ssMax, pvt  ' Initialize pending boundary stacks
    InitializeStacks twLb, twUb, twMax, pvt  ' Initialize pending runner stacks
    InitializeStacks lA_1, lA_2, ssBuf, cnt  ' Initialize working buffers

    lpL_1 = VarPtr(lA_1(n0))                 ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))                 ' Cache pointer to upper buffer
    lpStr = VarPtr(item)                     ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))  ' Cache pointer to the string array

    cnt = n0
    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA       ' Get pivot index position
        lItem = StringPtrs(pvt)              ' Grab current value into item
        CopyMemByR ByVal lpStr, lItem, n4

        For idx = lbA To pvt - n1
            If (StrComp(sA(idx), item, mMethod) = eComp) Then ' (idx > item)
                lA_2(ptr2) = StringPtrs(idx) ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = StringPtrs(idx) ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            If (StrComp(item, sA(idx), mMethod) = eComp) Then ' (idx < item)
                lA_1(ptr1) = StringPtrs(idx) ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = StringPtrs(idx) ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2 ©Rd-
        CopyMemByV Sum(lpS, lbA * n4), lpL_1, ptr1 * n4 ' 1 2 item 3 4
        StringPtrs(lbA + ptr1) = lItem       ' Re-assign current item
        CopyMemByV Sum(lpS, (lbA + ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If (ubA - lbA < n64) Then            ' Ignore false indicators
            If (inter1 = n0) Then            ' Reverse indicator
            ElseIf (ubA - lbA < n3) Then     ' Delegate to built-in Repeater on tiny chuncks
                walk = lbA
                Do Until walk = ubA
                    walk = walk + n1
                    CopyMemByV lpStr, Sum(lpS, walk * n4), n4 ' item = sA(walk)
                    find = walk
                    Do While StrComp(sA(find - n1), item, mMethod) = eComp
                        find = find - n1
                        If (find = lbA) Then Exit Do
                    Loop '-Repeater v45c ©Rd-
                    If (find < walk) Then
                        CopyMemByV Sum(lpS, (find + n1) * n4), Sum(lpS, find * n4), (walk - find) * n4
                        CopyMemByV Sum(lpS, find * n4), lpStr, n4 ' Move items up 1, sA(find) = item
                End If: Loop
                ptr1 = n0: ptr2 = n0
            End If
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then          ' Reverse
            ElseIf (ptr1 = n0) Then          ' Reverse Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                    ' Runners dislike super large ranges
                    CopyMemByR ByVal lpStr, StringPtrs(lbA + ptr1 - n1), n4
                    opt = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > opt                                         ' Runner do loop
                        If Not (StrComp(sA(run - n1), item, mMethod) = eComp) Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpS, run * n4), cast * n4        ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpS, run * n4), Sum(lpS, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpS, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    CopyMemByR ByVal lpStr, StringPtrs(base), n4
                    pvt = lbA + ptr1 + inter2
                    opt = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < opt                                         ' Runner do loop
                        If Not (StrComp(sA(run + n1), item, mMethod) = eComp) Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpS, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpS, (base + cast) * n4), Sum(lpS, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpS, base * n4), lpL_1, cast * n4       ' Re-assign items into position immediately above current item
            End If: End If: End If
        ElseIf (inter2 = n0) Then
            If (inter1 = ptr1) Then          ' Refresh
                ' Delegate to built-in Insert/Binary hybrid on ideal data state
                walk = lbA: mezz = ubA: idx = n0                              ' Initialize our walker variables
                opt = GetOptimalRange(ubA - lbA + n1)                         ' Get runners optimal range
                If opt > twMax Then InitializeStacks twLb, twUb, twMax, opt   ' Ensure enough stack space
                Do While walk < mezz ' ----==============================---- ' Do the twist while there's more items
                    walk = walk + n1                                          ' Walk up the array and use binary search to insert each item down into the sorted lower array
                    CopyMemByV lpStr, Sum(lpS, walk * n4), n4                 ' Grab current value into item
                    find = walk                                               ' Default to current position
                    ceil = walk - n1                                          ' Set ceiling to current position - 1
                    base = lbA                                                ' Set base to lower bound
                    Do While StrComp(sA(ceil), item, mMethod) = eComp   '  .  ' While current item must move down
                        midd = (base + ceil) \ n2                             ' Find mid point
                        Do Until StrComp(sA(midd), item, mMethod) = eComp     ' Step back up if below
                            base = midd + n1                                  ' Bring up the base
                            midd = (base + ceil) \ n2                         ' Find mid point
                            If midd = ceil Then Exit Do                       ' If we're up to ceiling
                        Loop                                                  ' Out of loop >= target pos
                        find = midd                                           ' Set provisional to new ceiling
                        If find = base Then Exit Do                           ' If we're down to base
                        ceil = midd - n1                                      ' Bring down the ceiling
                    Loop '-Twister v4 ©Rd-     .      . ...  .             .  ' Out of binary search loops
                    If (find < walk) Then                                     ' If current item needs to move down
                        CopyMemByV lpStr, Sum(lpS, find * n4), n4
                        run = walk + n1
                        Do Until run > mezz Or run - walk > opt               ' Runner do loop
                            If Not (StrComp(item, sA(run), mMethod) = eComp) Then Exit Do
                            run = run + 1
                        Loop: cast = (run - walk)
                        CopyMemByV lpL_2, Sum(lpS, walk * n4), cast * n4      ' Grab current value(s)
                        CopyMemByV Sum(lpS, (find + cast) * n4), Sum(lpS, find * n4), (walk - find) * n4 ' Move up items
                        CopyMemByV Sum(lpS, find * n4), lpL_2, cast * n4      ' Re-assign current value(s) into found pos
                        If cast > n1 Then
                            If Not run > mezz Then
                                idx = idx + n1
                                twLb(idx) = run - n1
                                twUb(idx) = mezz
                            End If
                            walk = find
                            mezz = find + cast - n1
                    End If: End If
                    If walk = mezz Then
                        If idx Then
                            walk = twLb(idx)
                            mezz = twUb(idx)
                            idx = idx - n1
                End If: End If: Loop     ' Out of walker do loop
                ' ----==========----
                ptr1 = n0: ptr2 = n0
            ElseIf (ptr2 = n0) Then      ' Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                    ' Runners dislike super large ranges
                    CopyMemByR ByVal lpStr, StringPtrs(lbA + ptr1 - n1), n4
                    opt = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > opt                                         ' Runner do loop
                        If Not (StrComp(sA(run - n1), item, mMethod) = eComp) Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpS, run * n4), cast * n4        ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpS, run * n4), Sum(lpS, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpS, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    CopyMemByR ByVal lpStr, StringPtrs(base), n4
                    pvt = lbA + ptr1 + inter2
                    opt = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < opt                                         ' Runner do loop
                        If Not (StrComp(sA(run + n1), item, mMethod) = eComp) Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpS, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpS, (base + cast) * n4), Sum(lpS, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpS, base * n4), lpL_1, cast * n4       ' Re-assign items into position immediately above current item
        End If: End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: ssLb(cnt) = lbA + ptr1 + n1: ssUb(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If cnt = n0 Then Exit Do
            lbA = ssLb(cnt): ubA = ssUb(cnt): cnt = cnt - n1
        End If
    Loop
    CopyMemByR ByVal lpSA, 0&, n4  ' De-reference our pointer to safearray header
    CopyMemByR ByVal lpStr, 0&, n4 ' De-reference our pointer to item variable
End Sub

' + Stable QuickSort v2.3 Indexed +++++++++++++++++++

' This is an indexed stable non-recursive quicksort.

' It uses a long array that holds references to the string arrays
' indices. This is known as an indexed sort. No changes are made
' to the source string array. This also allows the index array to
' be passed on to other sort processes to be further manipulated.

' After a sort procedure is run the long array is ready as a sorted
' index (lookup table) to the string array items.

' E.G sA(idxA(lo)) returns the lo item in the string array whose
' index may be anywhere in the string array.

Sub strStableSort2Indexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my indexed stable non-recursive quicksort
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim item As String, lpStr As Long, lpS As Long
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, run As Long, cast As Long
    Dim idx As Long, opt As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim ceil As Long, mezz As Long, lpB As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim idxItem As Long, lpI As Long
    Dim eComp As eSortOrder

    cnt = ubA - lbA + n1                ' Grab array item count
    If (cnt < n2) Then Exit Sub         ' If nothing to do then exit
    eComp = SortOrder                   ' Initialize compare variable
    pvt = (cnt \ n64) + n32             ' Allow for worst case senario + some

    ValidateIndexArray idxA, lbA, ubA         ' Validate the index array
    InitializeStacks ssLb, ssUb, ssMax, pvt   ' Initialize pending boundary stacks
    InitializeStacks twLb, twUb, twMax, pvt   ' Initialize pending runner stacks
    InitializeStacks lA_1, lA_2, ssBuf, cnt   ' Initialize working buffers

    lpL_1 = VarPtr(lA_1(n0))                  ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))                  ' Cache pointer to upper buffer
    lpStr = VarPtr(item)                      ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))   ' Cache pointer to the string array
    lpI = Sum(VarPtr(idxA(lbA)), -(lbA * n4)) ' Cache pointer to the index array

    cnt = n0
    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA   ' Get pivot index position
        idxItem = idxA(pvt)              ' Grab current value into item
        CopyMemByV lpStr, Sum(lpS, idxItem * n4), n4

        For idx = lbA To pvt - n1
            If (StrComp(sA(idxA(idx)), item, mMethod) = eComp) Then ' (idx > item)
                lA_2(ptr2) = idxA(idx)   ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(idx)   ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            If (StrComp(item, sA(idxA(idx)), mMethod) = eComp) Then ' (idx < item)
                lA_1(ptr1) = idxA(idx)   ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(idx)   ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2i ©Rd-
        lpB = VarPtr(idxA(lbA))          ' Cache pointer to current lb
        CopyMemByV lpB, lpL_1, ptr1 * n4
        idxA(lbA + ptr1) = idxItem       ' 1 2 item 3 4
        CopyMemByV Sum(lpB, (ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If (ubA - lbA < n64) Then        ' Ignore false indicators
            If (inter2 = ptr2) Then      ' Reverse indicator
            ElseIf (ubA - lbA < n3) Then ' Delegate to built-in Repeater on tiny chunks
                For walk = lbA + n1 To ubA
                   idxItem = idxA(walk)  ' Grab current value
                   CopyMemByV lpStr, Sum(lpS, idxItem * n4), n4 ' item = sA(walk)
                   find = walk
                   Do While StrComp(sA(idxA(find - n1)), item, mMethod) = eComp
                       find = find - n1
                       If (find = lbA) Then Exit Do
                   Loop '-Repeater v45i ©Rd-
                   If (find < walk) Then    ' Move items up 1, sA(find) = item
                       CopyMemByV Sum(lpI, (find + n1) * n4), Sum(lpI, find * n4), (walk - find) * n4
                       idxA(find) = idxItem ' Re-assign current item index into found pos
                End If: Next
                ptr1 = n0: ptr2 = n0
            End If
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then      ' Reverse
            ElseIf (ptr1 = n0) Then      ' Reverse Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                  ' Runners dislike super large ranges
                    CopyMemByV lpStr, Sum(lpS, idxA(lbA + ptr1 - n1) * n4), n4
                    opt = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > opt                                       ' Runner do loop
                        If Not (StrComp(sA(idxA(run - n1)), item, mMethod) = eComp) Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, run * n4), cast * n4      ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpI, run * n4), Sum(lpI, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpI, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    CopyMemByV lpStr, Sum(lpS, idxA(base) * n4), n4
                    pvt = lbA + ptr1 + inter2
                    opt = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < opt                                       ' Runner do loop
                        If Not (StrComp(sA(idxA(run + n1)), item, mMethod) = eComp) Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpI, (base + cast) * n4), Sum(lpI, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpI, base * n4), lpL_1, cast * n4     ' Re-assign items into position immediately above current item
            End If: End If: End If
        ElseIf (inter2 = n0) Then
            If (inter1 = ptr1) Then      ' Refresh
                ' Delegate to built-in Insert/Binary hybrid on ideal data state
                walk = lbA: mezz = ubA: idx = n0                                  ' Initialize our walker variables
                opt = GetOptimalRange(ubA - lbA + n1)                             ' Get runners optimal range
                If opt > twMax Then InitializeStacks twLb, twUb, twMax, opt       ' Ensure enough stack space
                Do While walk < mezz ' ----==================================---- ' Do the twist while there's more items
                    walk = walk + n1                                              ' Walk up the array and use binary search to insert each item down into the sorted lower array
                    CopyMemByV lpStr, Sum(lpS, idxA(walk) * n4), n4               ' Grab current value into item
                    find = walk                                                   ' Default to current position
                    ceil = walk - n1                                              ' Set ceiling to current position - 1
                    base = lbA                                                    ' Set base to lower bound
                    Do While StrComp(sA(idxA(ceil)), item, mMethod) = eComp '  .  ' While current item must move down
                        midd = (base + ceil) \ n2                                 ' Find mid point
                        Do Until StrComp(sA(idxA(midd)), item, mMethod) = eComp   ' Step back up if below
                            base = midd + n1                                      ' Bring up the base
                            midd = (base + ceil) \ n2                             ' Find mid point
                            If midd = ceil Then Exit Do                           ' If we're up to ceiling
                        Loop                                                      ' Out of loop >= target pos
                        find = midd                                               ' Set provisional to new ceiling
                        If find = base Then Exit Do                               ' If we're down to base
                        ceil = midd - n1                                          ' Bring down the ceiling
                    Loop '-Twister v4i ©Rd-    .       . ...   .               .  ' Out of binary search loops
                    If (find < walk) Then                                         ' If current item needs to move down
                        CopyMemByV lpStr, Sum(lpS, idxA(find) * n4), n4
                        run = walk + n1
                        Do Until run > mezz Or run - walk > opt                   ' Runner do loop
                            If Not (StrComp(item, sA(idxA(run)), mMethod) = eComp) Then Exit Do
                            run = run + 1
                        Loop: cast = (run - walk)
                        CopyMemByV lpL_2, Sum(lpI, walk * n4), cast * n4          ' Grab current value(s)
                        CopyMemByV Sum(lpI, (find + cast) * n4), Sum(lpI, find * n4), (walk - find) * n4 ' Move up items
                        CopyMemByV Sum(lpI, find * n4), lpL_2, cast * n4          ' Re-assign current value(s) into found pos
                        If cast > n1 Then
                            If Not run > mezz Then
                                idx = idx + n1
                                twLb(idx) = run - n1
                                twUb(idx) = mezz
                            End If
                            walk = find
                            mezz = find + cast - n1
                    End If: End If
                    If walk = mezz Then
                        If idx Then
                            walk = twLb(idx)
                            mezz = twUb(idx)
                            idx = idx - n1
                End If: End If: Loop     ' Out of walker do loop
                ' ----=================----
                ptr1 = n0: ptr2 = n0
            ElseIf (ptr2 = n0) Then      ' Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                  ' Runners dislike super large ranges
                    CopyMemByV lpStr, Sum(lpS, idxA(lbA + ptr1 - n1) * n4), n4
                    opt = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > opt                                       ' Runner do loop
                        If Not (StrComp(sA(idxA(run - n1)), item, mMethod) = eComp) Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, run * n4), cast * n4      ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpI, run * n4), Sum(lpI, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpI, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    CopyMemByV lpStr, Sum(lpS, idxA(base) * n4), n4
                    pvt = lbA + ptr1 + inter2
                    opt = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < opt                                       ' Runner do loop
                        If Not (StrComp(sA(idxA(run + n1)), item, mMethod) = eComp) Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpI, (base + cast) * n4), Sum(lpI, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpI, base * n4), lpL_1, cast * n4     ' Re-assign items into position immediately above current item
        End If: End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: ssLb(cnt) = lbA + ptr1 + n1: ssUb(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = ssLb(cnt): ubA = ssUb(cnt): cnt = cnt - n1
        End If
    Loop: CopyMemByR ByVal lpStr, 0&, n4 ' De-reference pointer to item variable
End Sub

' + Stable Presorter ++++++++++++++++++++++++++++++++++

' The purpose of this stable routine is to pre-sort only.

' If the data is already sorted it will complete only one pass and then exit.
' On the other hand, if the data is reverse-sorted or unsorted it will continue
' until the data is almost fully sorted and then exit.

' The benefit of this approach allows other routines that are optimized for pre-
' sorted data to complete the operation. Note that this pre-sorter is optimized
' for the Twister algorithm below and may need modification for other finishers.

' This is similar to the quicksort technique of dropping out of the sort operation
' before the data is fully sorted to reduce the number of recursions/iterations on
' small ranges of the data and to complete the job with a finisher such as an
' insertion sort.

' The only unique difference is this routine is stable and should therefore
' be complemented by a stable finisher.

Sub strPreSortStable(sA() As String, ByVal lbA As Long, ByVal ubA As Long) '-©Rd 2006-
    ' This is my stable non-recursive Avalanche based pre-sorter
    ' It dynamically sets minrange to pre-sort down to that range
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim item As String, lpStr As Long, lpS As Long
    Dim idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim base As Long, run As Long, cast As Long
    Dim bRevPretty As Boolean, minrange As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim lItem As Long, lpSA As Long
    Dim eComp As eSortOrder

    cnt = ubA - lbA + n1                ' Grab array item count
    If (cnt < n2) Then Exit Sub         ' If nothing to do then exit
    eComp = SortOrder                   ' Initialize compare variable
    pvt = (cnt \ n64) + n32             ' Allow for worst case senario + some

    lpSA = SubstArrayHeader(sA, lbA, ubA)
    InitializeStacks psLb, psUb, psMax, pvt  ' Initialize pending boundary stacks
    InitializeStacks lA_1, lA_2, ssBuf, cnt  ' Initialize working buffers

    lpL_1 = VarPtr(lA_1(n0))                 ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))                 ' Cache pointer to upper buffer
    lpStr = VarPtr(item)                     ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))  ' Cache pointer to the string array

    cnt = n0
    minrange = n1
 
    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA       ' Get pivot index position
        lItem = StringPtrs(pvt)              ' Grab current value into item
        CopyMemByR ByVal lpStr, lItem, n4

        For idx = lbA To pvt - n1
            If (StrComp(sA(idx), item, mMethod) = eComp) Then ' (idx > item)
                lA_2(ptr2) = StringPtrs(idx) ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = StringPtrs(idx) ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            If (StrComp(item, sA(idx), mMethod) = eComp) Then ' (idx < item)
                lA_1(ptr1) = StringPtrs(idx) ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = StringPtrs(idx) ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2 ©Rd-
        CopyMemByV Sum(lpS, lbA * n4), lpL_1, ptr1 * n4  ' 1 2 item 3 4
        StringPtrs(lbA + ptr1) = lItem       ' Re-assign current item
        CopyMemByV Sum(lpS, (lbA + ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If bRevPretty Then
        ElseIf (ubA - lbA < n32) Then  ' Ignore small ranges
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then    ' Reverse
                If (inter2) Then
                    minrange = ptr2 \ n16
                End If
            ElseIf (ptr1 = n0) Then    ' Reverse Pretty
                bRevPretty = True
                minrange = n1
            ElseIf (minrange < (ptr2 \ n3)) Then
                minrange = ptr2 \ n3   ' \3==iterate next (1/2 this size) but not again (1/4)
            End If
        ElseIf (inter2 = n0) Then
            If (inter1 = ptr1) Then    ' Resort
                minrange = ubA - lbA
            ElseIf (ptr2 = n0) Then    ' Pretty
                If (minrange < ptr1) Then
                    minrange = ptr1
                End If
            ElseIf (minrange < (ptr1 \ n3)) Then
                minrange = ptr1 \ n3
        End If: End If

        ' The following section has been added to handle a very hard job for a
        ' stable sorter - reverse-sorting case-insensitively on data pre-sorted
        ' case-sensitively. It utilises a runner technique to boost this very
        ' demanding operation - down from 2.0 to 1.5 seconds on 100,000 items.
        If bRevPretty Then
            If (ptr1 > inter1) And (inter1 < n50K) Then             ' Runners dislike super large ranges
                CopyMemByV lpStr, Sum(lpS, (lbA + ptr1 - n1) * n4), n4
                optimal = lbA + (inter1 \ n2)
                run = lbA + inter1
                Do While run > optimal                              ' Runner do loop
                    If Not (StrComp(sA(run - n1), item, mMethod) = eComp) Then Exit Do
                    run = run - n1
                Loop: cast = lbA + inter1 - run
                If cast Then
                    CopyMemByV lpL_1, Sum(lpS, run * n4), cast * n4 ' Grab items that stayed below current that should also be above items that have moved down below current
                    CopyMemByV Sum(lpS, run * n4), Sum(lpS, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                    CopyMemByV Sum(lpS, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                End If
            End If ' 1 2 1r item 4r 3 4
            If (inter2) And (ptr2 - inter2 < n50K) Then
                base = lbA + ptr1 + n1
                CopyMemByV lpStr, Sum(lpS, base * n4), n4
                pvt = lbA + ptr1 + inter2
                optimal = pvt + ((ptr2 - inter2) \ n2)
                run = pvt
                Do While run < optimal                              ' Runner do loop
                    If Not (StrComp(sA(run + n1), item, mMethod) = eComp) Then Exit Do
                    run = run + n1
                Loop: cast = run - pvt
                If cast Then
                    CopyMemByV lpL_1, Sum(lpS, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                    CopyMemByV Sum(lpS, (base + cast) * n4), Sum(lpS, base * n4), inter2 * n4 ' Move up items
                    CopyMemByV Sum(lpS, base * n4), lpL_1, cast * n4 ' Re-assign items into position immediately above current item
        End If: End If: End If

        If (ptr1 > minrange) Then
            If (ptr2 > minrange) Then cnt = cnt + n1: psLb(cnt) = lbA + ptr1 + n1: psUb(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > minrange) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = psLb(cnt): ubA = psUb(cnt): cnt = cnt - n1
        End If
    Loop
    CopyMemByR ByVal lpSA, 0&, n4  ' De-reference our pointer to safearray header
    CopyMemByR ByVal lpStr, 0&, n4 ' De-reference our pointer to item variable
End Sub

Sub strPreSortStableIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my indexed stable non-recursive Avalanche based pre-sorter
    ' It dynamically sets minrange to pre-sort down to that range
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim item As String, lpStr As Long, lpS As Long
    Dim idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim base As Long, run As Long, cast As Long
    Dim bRevPretty As Boolean, minrange As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim lpI As Long, lpB As Long
    Dim eComp As eSortOrder
    Dim idxItem As Long

    cnt = ubA - lbA + n1                  ' Grab array item count
    If (cnt < n2) Then Exit Sub           ' If nothing to do then exit
    eComp = SortOrder                     ' Initialize compare variable
    pvt = (cnt \ n64) + n32               ' Allow for worst case senario + some

    ValidateIndexArray idxA, lbA, ubA         ' Validate the index array
    InitializeStacks psLb, psUb, psMax, pvt   ' Initialize pending boundary stacks
    InitializeStacks lA_1, lA_2, ssBuf, cnt   ' Initialize working buffers

    lpL_1 = VarPtr(lA_1(n0))                  ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))                  ' Cache pointer to upper buffer
    lpStr = VarPtr(item)                      ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))   ' Cache pointer to the string array
    lpI = Sum(VarPtr(idxA(lbA)), -(lbA * n4)) ' Cache pointer to the index array

    cnt = n0
    minrange = n1

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA    ' Get pivot index position
        idxItem = idxA(pvt)               ' Grab current item index
        CopyMemByV lpStr, Sum(lpS, idxItem * n4), n4 ' Grab current value into item

        For idx = lbA To pvt - n1
            If (StrComp(sA(idxA(idx)), item, mMethod) = eComp) Then ' (idx > item)
                lA_2(ptr2) = idxA(idx)    ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(idx)    ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            If (StrComp(item, sA(idxA(idx)), mMethod) = eComp) Then ' (idx < item)
                lA_1(ptr1) = idxA(idx)    ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(idx)    ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2i ©Rd-
        lpB = VarPtr(idxA(lbA))           ' Cache pointer to current lo
        CopyMemByV lpB, lpL_1, ptr1 * n4  ' 1 2 item 3 4
        idxA(lbA + ptr1) = idxItem        ' Re-assign current
        CopyMemByV Sum(lpB, (ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If bRevPretty Then
        ElseIf (ubA - lbA < n32) Then     ' Ignore small ranges
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then       ' Reverse
                If (inter2) Then
                    minrange = ptr2 \ n16
                End If
            ElseIf (ptr1 = n0) Then       ' Reverse Pretty
                bRevPretty = True
                minrange = n1
            ElseIf (minrange < (ptr2 \ n3)) Then
                minrange = ptr2 \ n3      ' \3==iterate next (1/2 this size) but not again (1/4)
            End If
        ElseIf (inter2 = n0) Then
            If (inter1 = ptr1) Then       ' Resort
                minrange = ubA - lbA
            ElseIf (ptr2 = n0) Then       ' Pretty
                If (minrange < ptr1) Then
                    minrange = ptr1
                End If
            ElseIf (minrange < (ptr1 \ n3)) Then
                minrange = ptr1 \ n3
        End If: End If

        ' The following section has been added to handle a very hard job for a
        ' stable sorter - reverse-sorting case-insensitively on data pre-sorted
        ' case-sensitively. It utilises a runner technique to boost this very
        ' demanding operation - down from 2.0 to 1.5 seconds on 100,000 items.
        If bRevPretty Then
            If (ptr1 > inter1) And (inter1 < n50K) Then             ' Runners dislike super large ranges
                CopyMemByV lpStr, Sum(lpS, idxA(lbA + ptr1 - n1) * n4), n4
                optimal = lbA + (inter1 \ n2)
                run = lbA + inter1
                Do While run > optimal                              ' Runner do loop
                    If Not (StrComp(sA(idxA(run - n1)), item, mMethod) = eComp) Then Exit Do
                    run = run - n1
                Loop: cast = lbA + inter1 - run
                If cast Then
                    CopyMemByV lpL_1, Sum(lpI, run * n4), cast * n4 ' Grab items that stayed below current that should also be above items that have moved down below current
                    CopyMemByV Sum(lpI, run * n4), Sum(lpI, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                    CopyMemByV Sum(lpI, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                End If
            End If ' 1 2 1r item 4r 3 4
            If (inter2) And (ptr2 - inter2 < n50K) Then
                base = lbA + ptr1 + n1
                CopyMemByV lpStr, Sum(lpS, idxA(base) * n4), n4
                pvt = lbA + ptr1 + inter2
                optimal = pvt + ((ptr2 - inter2) \ n2)
                run = pvt
                Do While run < optimal                              ' Runner do loop
                    If Not (StrComp(sA(idxA(run + n1)), item, mMethod) = eComp) Then Exit Do
                    run = run + n1
                Loop: cast = run - pvt
                If cast Then
                    CopyMemByV lpL_1, Sum(lpI, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                    CopyMemByV Sum(lpI, (base + cast) * n4), Sum(lpI, base * n4), inter2 * n4 ' Move up items
                    CopyMemByV Sum(lpI, base * n4), lpL_1, cast * n4 ' Re-assign items into position immediately above current item
        End If: End If: End If

        If (ptr1 > minrange) Then
            If (ptr2 > minrange) Then cnt = cnt + n1: psLb(cnt) = lbA + ptr1 + n1: psUb(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > minrange) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = psLb(cnt): ubA = psUb(cnt): cnt = cnt - n1
        End If
    Loop: CopyMemByR ByVal lpStr, 0&, n4 ' De-reference our pointer to item variable
End Sub

' + Stable Insert/Binary Hybrid Twister v5 ++++++++++++++++

' This is the latest implementation of the Twister© algorithm.

' Just like previous versions of the Twister this sub-routine uses a unique
' binary search for fast insertion. But where they were based on an array walker
' version 4 uses an array runner and is TWICE as fast as the previous version!

' This algorithm is my outright fastest algorithm for re-sorting string arrays
' that have been sorted or that are mostly already sorted, for arrays of all sizes.

' Version 5 uses an Avalanche based pre-sorter to boost operations on unsorted data,
' and so requires the PreSortStable function above. Note this version is identical
' in code to version 4 but calls out to the pre-sorter hence the v5 designation.

Sub strTwist5(sA() As String, ByVal lbA As Long, ByVal ubA As Long)
    ' Optimized Insert/Binary hybrid using CopyMemory
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, ceil As Long, item As String
    Dim cast As Long, mezz As Long, run As Long
    Dim lpStr As Long, lpS As Long, lpB As Long
    Dim idx As Long, optimal As Long
    Dim eComp As eCompareResult
    walk = (ubA - lbA) + n1      ' Grab array item count
    If walk < n2 Then Exit Sub   ' If nothing to do then exit
    strPreSortStable sA, lbA, ubA
    eComp = SortOrder                                             ' Initialize compare variable
    optimal = GetOptimalRange(walk, True)                         ' Limit runners range
    InitializeStacks twLb, twUb, twMax, optimal                   ' Initialize pending boundary stacks
    InitializeBuffer twBuf, twBufMax, walk                        ' Initialize working buffer
    lpStr = VarPtr(item)                                          ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))                       ' Cache pointer to the array
    lpB = VarPtr(twBuf(n0))                                       ' Cache pointer to the buffer
    walk = lbA: mezz = ubA                                        ' Initialize our walker variables
    Do While walk < mezz ' ----==============================---- ' Do the twist while there's more items
        walk = walk + n1                                          ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, Sum(lpS, walk * n4), n4                 ' Grab current value into item
        find = walk                                               ' Default to current position
        ceil = walk - n1                                          ' Set ceiling to current position - 1
        base = lbA                                                ' Set base to lower bound
        Do While StrComp(sA(ceil), item, mMethod) = eComp   '  .  ' While current item must move down
            midd = (base + ceil) \ n2                             ' Find mid point
            Do Until StrComp(sA(midd), item, mMethod) = eComp     ' Step back up if below
                base = midd + n1                                  ' Bring up the base
                midd = (base + ceil) \ n2                         ' Find mid point
                If midd = ceil Then Exit Do                       ' If we're up to ceiling
            Loop                                                  ' Out of loop >= target pos
            find = midd                                           ' Set provisional to new ceiling
            If find = base Then Exit Do                           ' If we're down to base
            ceil = midd - n1                                      ' Bring down the ceiling
        Loop '-Twister v5 ©Rd-     .      . ...  .             .  ' Out of binary search loops
        If (find < walk) Then                                     ' If current item needs to move down
            CopyMemByV lpStr, Sum(lpS, find * n4), n4
            run = walk + n1
            Do Until run > mezz Or run - walk > optimal           ' Runner do loop
                If Not (StrComp(item, sA(run), mMethod) = eComp) Then Exit Do
                run = run + 1
            Loop: cast = (run - walk)
            CopyMemByV lpB, Sum(lpS, walk * n4), cast * n4        ' Grab current value(s)
            CopyMemByV Sum(lpS, (find + cast) * n4), Sum(lpS, find * n4), (walk - find) * n4 ' Move up items
            CopyMemByV Sum(lpS, find * n4), lpB, cast * n4        ' Re-assign current value(s) into found pos
            If cast > n1 Then
                If Not run > mezz Then
                    idx = idx + n1
                    twLb(idx) = run - n1
                    twUb(idx) = mezz
                End If
                walk = find
                mezz = find + cast - n1
        End If: End If
        If walk = mezz Then
            If idx Then
                walk = twLb(idx)
                mezz = twUb(idx)
                idx = idx - n1
    End If: End If: Loop           ' Out of walker do loop
    CopyMemByR ByVal lpStr, 0&, n4 ' De-reference pointer to item variable
    ' ----=================----
End Sub

Sub strTwist5Indexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' Optimized Indexed Insert/Binary hybrid using CopyMemory
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, ceil As Long, item As String
    Dim cast As Long, mezz As Long, run As Long
    Dim lpStr As Long, lpS As Long, lpB As Long
    Dim idx As Long, optimal As Long
    Dim eComp As eCompareResult
    Dim lpI As Long
    walk = (ubA - lbA) + n1             ' Grab array item count
    If walk < n2 Then Exit Sub          ' If nothing to do then exit
    strPreSortStableIndexed sA, idxA, lbA, ubA
    eComp = SortOrder                                                 ' Initialize compare variable
    optimal = GetOptimalRange(walk, True)                             ' Limit runners range
    InitializeStacks twLb, twUb, twMax, optimal                       ' Initialize pending boundary stacks
    InitializeBuffer twBuf, twBufMax, walk                            ' Initialize working buffer
    lpStr = VarPtr(item)                                              ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * n4))                           ' Cache pointer to the array
    lpI = Sum(VarPtr(idxA(lbA)), -(lbA * n4))                         ' Cache pointer to the index array
    lpB = VarPtr(twBuf(n0))                                           ' Cache pointer to the buffer
    walk = lbA: mezz = ubA                                            ' Initialize our walker variables
    Do While walk < mezz ' ----==================================---- ' Do the twist while there's more items
        walk = walk + n1                                              ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, Sum(lpS, idxA(walk) * n4), n4               ' Grab current value into item
        find = walk                                                   ' Default to current position
        ceil = walk - n1                                              ' Set ceiling to current position - 1
        base = lbA                                                    ' Set base to lower bound
        Do While StrComp(sA(idxA(ceil)), item, mMethod) = eComp '  .  ' While current item must move down
            midd = (base + ceil) \ n2                                 ' Find mid point
            Do Until StrComp(sA(idxA(midd)), item, mMethod) = eComp   ' Step back up if below
                base = midd + n1                                      ' Bring up the base
                midd = (base + ceil) \ n2                             ' Find mid point
                If midd = ceil Then Exit Do                           ' If we're up to ceiling
            Loop                                                      ' Out of loop >= target pos
            find = midd                                               ' Set provisional to new ceiling
            If find = base Then Exit Do                               ' If we're down to base
            ceil = midd - n1                                          ' Bring down the ceiling
        Loop '-Twister v5i ©Rd-    .       . ...   .               .  ' Out of binary search loops
        If (find < walk) Then                                         ' If current item needs to move down
            CopyMemByV lpStr, Sum(lpS, idxA(find) * n4), n4
            run = walk + n1
            Do Until run > mezz Or run - walk > optimal               ' Runner do loop
                If Not (StrComp(item, sA(idxA(run)), mMethod) = eComp) Then Exit Do
                run = run + 1
            Loop: cast = (run - walk)
            CopyMemByV lpB, Sum(lpI, walk * n4), cast * n4            ' Grab current value(s)
            CopyMemByV Sum(lpI, (find + cast) * n4), Sum(lpI, find * n4), (walk - find) * n4 ' Move up items
            CopyMemByV Sum(lpI, find * n4), lpB, cast * n4            ' Re-assign current value(s) into found pos
            If cast > n1 Then
                If Not run > mezz Then
                    idx = idx + n1
                    twLb(idx) = run - n1
                    twUb(idx) = mezz
                End If
                walk = find
                mezz = find + cast - n1
        End If: End If
        If walk = mezz Then
            If idx Then
                walk = twLb(idx)
                mezz = twUb(idx)
                idx = idx - n1
    End If: End If: Loop           ' Out of walker do loop
    CopyMemByR ByVal lpStr, 0&, n4 ' De-reference pointer to item variable
    ' ----=================----
End Sub

' + ArrayPtr +++++++++++++++++++++++++++++++++++++++++++++++++

' This function returns a pointer to the SAFEARRAY header of
' any Visual Basic array, including a Visual Basic string array.

' Substitutes both ArrPtr and StrArrPtr.

' This function will work with vb5 or vb6 without modification.

Function ArrayPtr(Arr) As Long
    Dim iDataType As Integer
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, n2                           ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then               ' if a valid array was passed
        CopyMemByR ArrayPtr, ByVal Sum(VarPtr(Arr), n8), n4 ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
UnInit:
End Function

' + Validate Index Array +++++++++++++++++++++++++++++++++++++

' This will prepare the passed index array if it is not already.

' This sub-routine determines if the index array passed is either:
' [A] uninitialized or Erased
'     initialized to invalid boundaries
'     initialized to valid boundaries but not prepared
' [B] initialized to extended boundaries and not fully prepared
' [C] prepared for the sort process by the For loop
'     has been modified by a previous sort process

' If the condition is determined to be [A] then it is prepared by
' executing the For loop code, if the condition is determined to
' be [B] then it is prepared only from the old ub to the new ub,
' otherwise if [C] nothing is done.

' This permits subsequent sorting of the data without interfering
' with the index array if it is already sorted (based on criteria
' that may differ from the current process, for example, or some
' items have been modified in the sorted array).

' It also permits refresh-sorting of data that has additional
' items added to the top of the sorted array without interfering
' with the index array and so does not require a full resort.

' Otherwise, it ensures that the index array is in the required
' pre-sort state produced by the For loop.

Sub ValidateIndexArray(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim bReDim As Boolean, bReDimPres As Boolean
    Dim lb As Long, ub As Long, j As Long
    lb = &H80000000: ub = &H7FFFFFFF
    bReDim = Not InitedArray(idxA, lb, ub)
    If bReDim = False Then
        bReDim = (lbA < lb)
        bReDimPres = (ubA > ub)
    End If '-©Rd-
    If bReDim Then
        ReDim idxA(lbA To ubA) As Long
    ElseIf bReDimPres Then
        ReDim Preserve idxA(lb To ubA) As Long
    End If
    If (idxA(ubA) = n0) Then
        If (idxA(lbA) = n0) Then
            For j = lbA To ubA
                idxA(j) = j
            Next
        ElseIf bReDimPres Then
            For j = ub + n1 To ubA
                idxA(j) = j
            Next
        End If
    End If
End Sub

' + Inited Array +++++++++++++++++++++++++++++++++++++++++++++

' This function determines if the passed array is initialized,
' and if so will return -1.

' It will also optionally indicate whether the array can be redimmed;
' in which case it will return -2.

' If the array is uninitialized (never redimmed or has been erased)
' it will return 0 (zero).

Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant headache,
    ' and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY, lpSA As Long
    Dim iDataType As Integer, lOffset As Long
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, n2                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemByR lpSA, ByVal Sum(VarPtr(Arr), n8), n4 ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then                ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case lpSA already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMemByR lpSA, ByVal lpSA, n4             ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = (lpSA <> n0)
        If InitedArray Then
            CopyMemByR tSA.cDims, ByVal lpSA, n4
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + ((tSA.fFeatures And FADF_FIXEDSIZE) <> FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = n16 + ((tSA.cDims - n1) * n8)
            CopyMemByR tSA.cElements, ByVal Sum(lpSA, lOffset), n8
            tSA.lUBound = tSA.lLBound + tSA.cElements - n1
            If (lbA < tSA.lLBound) Then lbA = tSA.lLBound
            If (ubA > tSA.lUBound) Then ubA = tSA.lUBound
    End If: End If
UnInit:
End Function

' + Sum ++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Enables valid addition and subtraction of unsigned long ints.
' Treats lPtr as an unsigned long and returns an unsigned long.
' Allows safe arithmetic operations on memory address pointers.
' Assumes valid pointer and pointer offset.

Private Function Sum(ByVal lPtr As Long, ByVal lOffset As Long) As Long
    If lOffset > n0 Then
        If lPtr And DW_MSB Then ' if ptr < 0
           Sum = lPtr + lOffset ' ignors > unsigned int max
        ElseIf (lPtr Or DW_MSB) < -lOffset Then
           Sum = lPtr + lOffset ' result is below signed int max
        Else                    ' result wraps to min signed int
           Sum = (lPtr + DW_MSB) + (lOffset + DW_MSB)
        End If
    ElseIf lOffset = n0 Then
        Sum = lPtr
    Else 'If lOffset < 0 Then
        If (lPtr And DW_MSB) = n0 Then ' if ptr > 0
           Sum = lPtr + lOffset ' ignors unsigned int < zero
        ElseIf (lPtr - DW_MSB) >= -lOffset Then
           Sum = lPtr + lOffset ' result is above signed int min
        Else                    ' result wraps to max signed int
           Sum = (lOffset - DW_MSB) + (lPtr - DW_MSB)
        End If
    End If
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Function SubstArrayHeader(sA() As String, ByVal lbA As Long, ByVal ubA As Long) As Long
    With StringPtrs_Header
        .cDims = n1                ' 1 Dimensional
        .fFeatures = FADF_NO_REDIM ' Cannot REDIM the array
        .cbElements = n4           ' This is a long array
        .cLocks = n1               ' Lock the array
        .pvData = VarPtr(sA(lbA))  ' Get the ptr to the first String descriptor
        .cElements = ubA - lbA + n1
        .lLBound = lbA
    End With
    SubstArrayHeader = ArrayPtr(StringPtrs)
    CopyMemByR ByVal SubstArrayHeader, VarPtr(StringPtrs_Header), n4
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub InitializeStacks(LBstack() As Long, UBstack() As Long, ByRef pCurMax As Long, ByVal NewMax As Long)
    If NewMax > pCurMax Then
        ReDim LBstack(n0 To NewMax) As Long   ' Stack to hold pending lower boundries
        ReDim UBstack(n0 To NewMax) As Long   ' Stack to hold pending upper boundries
        pCurMax = NewMax
    End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub InitializeBuffer(buffer() As Long, ByRef pCurMax As Long, ByVal NewMax As Long)
    If NewMax > pCurMax Then
        ReDim buffer(n0 To NewMax) As Long
        pCurMax = NewMax
    End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Function GetOptimalRange(ByVal lCount As Long, Optional ByVal fOptimal As Boolean) As Long
    Dim optimal As Long, range As Single ' CraZy performance curve
    If lCount > n20K Then optimal = n12 * (lCount \ n10K - n2)
    If fOptimal Then range = rRunner5 Else range = rRunner4
    GetOptimalRange = (lCount * range) - optimal + n4
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function strVerifySort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As eCompareMethod, Optional ByVal Order As eSortOrder = Default) As Boolean
    On Error GoTo FreakOut
    Dim walk As Long
    For walk = lbA + n1 To ubA
        If StrComp(sA(walk - n1), sA(walk), CaseSensitive) = Order Then Exit Function
    Next
FreakOut:
    strVerifySort = (walk > ubA)
End Function

Function strVerifyIndexed(sA() As String, lA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As eCompareMethod, Optional ByVal Order As eSortOrder = Default) As Boolean
    On Error GoTo FreakOut
    Dim walk As Long
    For walk = lbA + n1 To ubA
        If StrComp(sA(lA(walk - n1)), sA(lA(walk)), CaseSensitive) = Order Then Exit Function
    Next
FreakOut:
    strVerifyIndexed = (walk > ubA)
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!
