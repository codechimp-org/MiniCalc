'========================================================
'
'	Class Name:  SubclassedSystemMenu
'	Description: Object that allows a modified system menu 
'                to be implemented and interacted with. The 
'                object is designed to add a new seperator and 
'                "About..." menu item to a windows system menu
'
'========================================================

Public Class SubclassedSystemMenu
    Inherits System.Windows.Forms.NativeWindow
    Implements IDisposable

#Region "Win32 API Declares"
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Int32, _
                                                         ByVal bRevert As Boolean) As Int32

    Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Int32, _
                                                                          ByVal wFlags As Int32, _
                                                                          ByVal wIDNewItem As Int32, _
                                                                          ByVal lpNewItem As String) As Int32

    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Int32, ByVal pos As Int32, ByVal flags As Int32) As Int32

    Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Int32, ByVal pos As Int32) As Int32

    Private Declare Function ModifyMenu Lib "user32" (ByVal hMenu As Int32, ByVal pos As Int32, ByVal flags As Int32, ByVal newPos As Integer, ByVal lpNewItem As String) As Int32

    Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Int32, ByVal pos As Int32, ByVal flags As Int32) As Int32

    Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Int32, ByVal uItem As Int32, ByVal fByPosition As Boolean, ByRef lpmii As MENUITEMINFO) As Boolean

    Private Declare Function SetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Int32, ByVal uItem As Int32, ByVal fByPosition As Boolean, ByRef lpmii As MENUITEMINFO) As Boolean
#End Region

#Region "Constants"
    Private Const MF_STRING As Int32 = &H0       ' Menu string format
    Private Const MF_SEPARATOR As Int32 = &H800  ' Menu separator
    Private Const MF_BYPOSITION As Int32 = &H400
    Private Const MF_REMOVE As Int32 = &H1000
    Private Const MF_CHECKED As Int32 = &H8
    Private Const MF_APPEND As Int32 = &H100
    Private Const MF_GRAYED As Int32 = &H1
    Private Const MF_DISABLED As Int32 = &H2
    Private Const MF_BITMAP As Int32 = &H4
    Private Const MF_RADIOCHECK As Int32 = &H200
    Private Const MF_BREAK As Int32 = &H40
    Private Const MF_BARBREAK As Int32 = &H20
    Private Const WM_SYSCOMMAND As Int32 = &H112 ' System menu 
#End Region

#Region "Structures"
    Public Structure MENUITEMINFO
        Public cbSize As Integer
        Public fMask As Integer
        Public fType As Integer
        Public fState As Integer
        Public wID As Integer
        Public hSubMenu As IntPtr
        Public hbmpChecked As IntPtr
        Public hbmpUnchecked As IntPtr
        Public dwItemData As IntPtr
        Public dwTypeData As String
        Public cch As Integer
        Public hbmpItem As IntPtr
    End Structure
#End Region

#Region "Member Variables"
    Private mintSystemMenu As Int32 = 0                 ' Parent system menu handle
    Private mintHandle As Int32 = 0                     ' Local parent window handle    
#End Region

#Region "Events"
    Public Event SysMenuClick(ByVal menuItemID As Integer)
#End Region

#Region "Properties"
    Public ReadOnly Property SysMenuHandle() As Integer
        Get
            Return mintSystemMenu
        End Get
    End Property
#End Region

#Region "Constructor"
    '========================================================
    '
    '   Method Name:        New
    '	Description:	    Constructor. Creates menu items and assigns subclass
    '
    '   Inputs:             intWindowHandle : Parent window handle for message 
    '                                         subclass and adding new menu items 
    '                                         to parent system menu
    '
    '   Return Value:       None
    '
    '========================================================
    Public Sub New(ByVal intWindowHandle As Int32)

        Me.AssignHandle(New IntPtr(intWindowHandle))

        mintHandle = intWindowHandle

        ' Retrieve the system menu handle
        mintSystemMenu = GetSystemMenu(mintHandle, False)

    End Sub
#End Region

#Region "Methods"

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case m.Msg
            Case WM_SYSCOMMAND

                MyBase.WndProc(m)
                RaiseEvent SysMenuClick(m.WParam.ToInt32)

            Case Else
                MyBase.WndProc(m)
        End Select

    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose

        If Not Me.Handle.Equals(IntPtr.Zero) Then
            Me.ReleaseHandle()
        End If

    End Sub

    Private Function AddNewSystemMenuItem(ByVal menuItemID As Integer, ByVal menuItemText As String) As Boolean
        Try
            ' Append the extra system menu items
            Return AppendToSystemMenu(mintSystemMenu, menuItemID, menuItemText)

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function AppendToSystemMenu(ByVal intHandle As Int32, _
                                        ByVal menuID As Integer, _
                                        ByVal strText As String) As Boolean

        Try
            ' Add the About... menu item
            Dim intRet As Int32 = AppendMenu(intHandle, MF_STRING, menuID, strText)

            If intRet = 1 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function AppendSeperatorToSystemMenu(ByVal intHandle As Int32, _
                                    ByVal menuID As Integer) As Boolean

        Try
            ' Add the seperator menu item
            Dim intRet As Int32 = AppendMenu(intHandle, MF_SEPARATOR, 0, String.Empty)

            If intRet = 1 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public WriteOnly Property Checked(ByVal intHandle As Int32, ByVal menuID As Integer) As Boolean
        'Get
        '    Dim mii As New MENUITEMINFO
        '    GetMenuItemInfo(intHandle, menuID, True, mii)

        '    Return mii.hbmpChecked.ToInt32 = 1
        'End Get
        Set(ByVal value As Boolean)
            If value = True Then
                CheckMenuItem(intHandle, menuID, MF_CHECKED)
            Else
                CheckMenuItem(intHandle, menuID, MF_STRING)
            End If
        End Set
    End Property

#End Region

End Class
