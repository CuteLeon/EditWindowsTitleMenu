Public Class Form1
    Declare Function GetSystemMenu Lib "user32" Alias "GetSystemMenu" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
    Declare Function GetMenuItemCount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Integer) As Integer
    Declare Function DrawMenuBar Lib "user32" Alias "DrawMenuBar" (ByVal hwnd As Integer) As Integer
    Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As String) As Integer
    Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As String) As Integer
    Declare Function RemoveMenu Lib "user32" Alias "RemoveMenu" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Declare Function DeleteMenu Lib "user32" Alias "DeleteMenu" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Declare Function EnableMenultem Lib "user32" Alias "EnableMenuItem" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Declare Function CheckMenuItem Lib "user32" Alias "CheckMenuItem" (ByVal hMenu As Integer, uIDCheckItem As Integer, uCheck As Integer) As Integer
    Declare Function CheckMenuRadioItem Lib "user32" Alias "CheckMenuRadioItem" (ByVal hMenu As Integer, idFirst As Integer, idLast As Integer, idCheck As Integer, uFlags As Integer) As Integer
    Declare Function GetMenuState Lib "user32" Alias "GetMenuState" (ByVal hMenu As Integer, nPosition As Integer, uFlags As Integer) As Integer

    Private Const SC_SIZE = &HF000
    Private Const SC_MOVE = &HF010
    Private Const SC_MINIMIZE = &HF020
    Private Const SC_MAXIMIZE = &HF030
    Private Const SC_NEXTWINDOW = &HF040
    Private Const SC_PREVWINDOW = &HF050
    Private Const SC_CLOSE = &HF060
    Private Const SC_VSCROLL = &HF070
    Private Const SC_HSCROLL = &HF080
    Private Const SC_MOUSEMENU = &HF090
    Private Const SC_KEYMENU = &HF100
    Private Const SC_ARRANGE = &HF110
    Private Const SC_RESTORE = &HF120
    Private Const SC_TASKLIST = &HF130
    Private Const SC_SCREENSAVE = &HF140
    Private Const SC_HOTKEY = &HF150


    Private Const MF_STRING = &H0L '指定菜单项是一个正文字符串。
    Private Const MF_BITMAP = &H4L '将一个位图用作菜单项。
    Private Const MF_OWNERDRAW = &H100L '指定该菜单项为自绘制菜单项。菜单第一次显示前，拥有菜单的窗口接收一个WM_MEASUREITEM消息来得到菜单项的宽和高。然后，只要菜单项被修改，都将发送WM_DRAWITEM消息给菜单拥有者的窗口程序。

    Private Const MF_BYCOMMAND = &H0L '表示参数uId给出菜单项的标识符。如果MF_BYCOMMAND和MF_BYPOSITION都没被指定，则MF_BYCOMMAND是缺省值。
    Private Const MF_BYPOSITION = &H400L '表示参数uId给出菜单项相对于零的位置。
    Private Const MF_SEPARATOR = &H800L '创建一个水平分隔线

    Private Const MF_ENABLED = &H0L '表明菜单项有效。
    Private Const MF_GRAYED = &H1L '使菜单项无效并变灰。
    Private Const MF_DISABLED = &H2L '使菜单项无效。

    Private Const MF_UNHILITE = &H0L '取消加亮菜单项。
    Private Const MF_HILITE = &H80L '加亮菜单项。

    Private Const MF_UNCHECKED = &H0L '取消放置于菜单项旁边的标记。
    Private Const MF_CHECKED = &H8L '放置选取标记于菜单项旁边。
    Private Const MF_USECHECKBITMAPS = &H200L

    Private Const MF_POPUP = &H10L '指定菜单打开一个下拉式菜单或子菜单。参数uIDNewltem下拉式菜单或子菜单的句柄。
    Private Const MF_MENUBREAK = &H20L '将菜单项放于新行（对菜单条）或无分隔列地放于新列（对下拉式菜单、子菜单或快捷菜单）。
    Private Const MF_MENUBARBREAK = &H40L '对下拉式菜单、子菜单和快捷菜单，新列和旧列由垂直线隔开，其余功能同MF_MENUBREAK标志。

    '下列标志组不能被一起使用：
    'MF_BYCOMMAND && MF_BYPOSITION
    'MF_DISABLED  && MF_ENABLED && MF_GRAYED
    'MF_BITMAP && MF_STRING && MF_OWNERDRAW && MF_SEPARATOR
    'MF_MENUBARBREAK && MF_MENUBREAK
    'MF_CHECKED && MF_UNCHECKED


    Private Const WM_SYSCOMMAND = &H112

    Private IsChange As Boolean


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ChangeSysMenu()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If (m.Msg = WM_SYSCOMMAND) Then
            Select Case m.WParam.ToInt32()

                Case &H1000
                    Dim SysMenu = GetSystemMenu(Me.Handle, False)
                    Dim IntA As Integer = GetMenuState(SysMenu, &H1000, MF_BYCOMMAND)
                    If (IntA And MF_CHECKED) Then
                        Dim Checked = CheckMenuItem(SysMenu, &H1000, MF_UNCHECKED)
                        Dim DisColse = EnableMenultem(SysMenu, SC_CLOSE, MF_BYCOMMAND Or MF_ENABLED)
                    Else
                        Dim Checked = CheckMenuItem(SysMenu, &H1000, MF_CHECKED)
                        Dim DisColse = EnableMenultem(SysMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)
                    End If
                Case &H1001
                    Dim SysMenu = GetSystemMenu(Me.Handle, True)
                    IsChange = False
                Case Else

            End Select

        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call ChangeSysMenu()
    End Sub

    Private Sub ChangeSysMenu()
        If IsChange Then
            MessageBox.Show("系统菜单菜单已经修改，若要查看当前系统菜单" & vbCrLf & vbCrLf & "请单击标题栏图标或在标题栏任意位置右键单击", "API控制系统菜单", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim SysMenu = GetSystemMenu(Me.Handle, False)
        Dim RemMax = RemoveMenu(SysMenu, SC_MAXIMIZE, MF_BYCOMMAND Or MF_GRAYED)
        Dim RemMin = RemoveMenu(SysMenu, SC_MINIMIZE, MF_BYCOMMAND Or MF_GRAYED)
        'Dim RemMov = RemoveMenu(SysMenu, SC_MOVE, MF_BYCOMMAND Or MF_GRAYED)’删除这个菜单项会导致无法拖动标题栏移动窗口
        Dim RemMSize = RemoveMenu(SysMenu, SC_SIZE, MF_BYCOMMAND Or MF_GRAYED)
        Dim RemRes = RemoveMenu(SysMenu, SC_RESTORE, MF_BYCOMMAND Or MF_GRAYED)

        InsertMenu(SysMenu, 0, MF_BYPOSITION, &H1000, "禁止关闭窗口")
        InsertMenu(SysMenu, 1, MF_BYPOSITION Or MF_SEPARATOR, 0, String.Empty)
        InsertMenu(SysMenu, 2, MF_BYPOSITION, &H1001, "还原系统菜单")
        DrawMenuBar(SysMenu)
        IsChange = True
        MessageBox.Show("系统菜单菜单已经修改，若要查看当前系统菜单" & vbCrLf & vbCrLf & "请单击标题栏图标或在标题栏任意位置右键单击", "API控制系统菜单", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
