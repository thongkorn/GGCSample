#Region "About"
' / -----------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Sample code for GridGroupingControl of Syncfusion Community.
' / Microsoft Visual Basic .NET (2010) & MS Access 2010.
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / -----------------------------------------------------------------
#End Region

'// Syncfusion
Imports Syncfusion.Windows.Forms
Imports Syncfusion.Windows.Forms.Grid
Imports Syncfusion.Grouping
Imports Syncfusion.Drawing
'// DataBase
Imports System.Data.OleDb

Public Class frmGGCSample

    Private Sub frmGGCSample_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '// Connect DataBase
        Call ConnectDataBase()
        '// Show all records.
        Call RetrieveData()
    End Sub

    ' / -----------------------------------------------------------------
    ' / blnSearch = True, It's search for specified data, False = all data is displayed.
    Private Sub RetrieveData(Optional ByVal blnSearch As Boolean = False)
        strSQL = _
            " SELECT Countries.CountryPK, Countries.A2, Countries.Country, Countries.Capital, " & _
            " Countries.Population, Zones.ZoneName " & _
            " FROM Countries INNER JOIN Zones ON Countries.ZoneFK = Zones.ZonePK "

        '// blnSearch = True for Search
        If blnSearch Then
            strSQL = strSQL & _
                " WHERE " & _
                " [A2] " & " Like '%" & txtSearch.Text & "%'" & " OR " & _
                " [Country] " & " Like '%" & txtSearch.Text & "%'" & " OR " & _
                " [Capital] " & " Like '%" & txtSearch.Text & "%'" & " OR " & _
                " [ZoneName] " & " Like '%" & txtSearch.Text & "%'" & _
                " ORDER BY Country "
        Else
            strSQL = strSQL & " ORDER BY Country "
        End If
        '//
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        '// Data Adapter. 
        DA = New OleDbDataAdapter(strSQL, Conn)
        ' Fill Data Set. 
        DS = New DataSet
        DA.Fill(DS)
        Me.GGC.DataSource = DS.Tables(0)
        lblRecordCount.Text = "[Total: " & Format(DS.Tables(0).Rows.Count, "#,##") & " Records.]"
        DA.Dispose()
        DS.Dispose()
        Conn.Close()
        '//
        Call InitGridGroup()
        '//
        txtSearch.Clear()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        '// Show all records.
        Call RetrieveData(False)
    End Sub

    Private Sub txtSearch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If Trim(txtSearch.Text) = "" Or Len(Trim(txtSearch.Text)) = 0 Then Exit Sub
        '// Undesirable characters for the database ex.  ', * or %
        txtSearch.Text = Replace(Trim(txtSearch.Text), "'", "").Replace("%", "").Replace("*", "")
        '/ RetrieveData(True) It means searching for information.
        If e.KeyChar = Chr(13) Then '// Press Enter
            '// No beep.
            e.Handled = True
            '// Search Data with True Parameter.
            Call RetrieveData(True)
        End If
    End Sub

    ' / -----------------------------------------------------------------
    ' / Initilized GridGroupingControl
    Private Sub InitGridGroup()
        '// Initialize Columns GridGroup
        With Me.GGC.TableDescriptor
            '// Hidden Primary Key Column
            .VisibleColumns.Remove("CountryPK")
            '/ Using Column Name
            .Columns("A2").HeaderText = "A2"
            .Columns("Country").HeaderText = "Country"
            .Columns("Capital").HeaderText = "Capital"
            '// Format Population
            With .Columns("Population")
                .HeaderText = "Population"
                .Appearance.AnyRecordFieldCell.CellValueType = GetType(Int32)
                .Appearance.AnyRecordFieldCell.Format = "N0"
            End With
            .Columns("ZoneName").HeaderText = "Zone Name"
        End With
        '// GridVerticalAlignment.Middle
        For i As Byte = 0 To 5
            With Me.GGC.TableDescriptor
                .Columns(i).Appearance.AnyRecordFieldCell.VerticalAlignment = GridVerticalAlignment.Middle
                .Columns(i).AllowGroupByColumn = False
                ' // Set Font any Columns.
                .Columns(i).Appearance.AnyRecordFieldCell.Font = New Syncfusion.Windows.Forms.Grid.GridFontInfo(New Font("Tahoma", 11.0F, FontStyle.Regular))
            End With
        Next

        '// Initialize normal GridGroupingControl
        With Me.GGC
            ' Allows GroupDropArea to be visible
            .ShowGroupDropArea = False  ' Disable
            '// Hidden Top Level of Grouping
            .TopLevelGroupOptions.ShowCaption = False
            '// Styles Theme
            '// .GridVisualStyles = Syncfusion.Windows.Forms.GridVisualStyles.SystemTheme
            '// Metro Styles
            .GridVisualStyles = Syncfusion.Windows.Forms.GridVisualStyles.Office2010Blue
            ' Disables editing in GridGroupingControl
            .ActivateCurrentCellBehavior = GridCellActivateAction.None
            ' Disables editing in GridGroupingControl
            .TableDescriptor.AllowNew = False
            '// Autofit Columns
            .AllowProportionalColumnSizing = True

            '// Row Height
            .Table.DefaultRecordRowHeight = 25
            '// 
            .Table.DefaultCaptionRowHeight = 25
            .Table.DefaultColumnHeaderRowHeight = 30    '// Columns Header Row Height

            '// Selection
            .TableOptions.ListBoxSelectionMode = SelectionMode.One
            'Selection Back color
            .TableOptions.SelectionBackColor = Color.Firebrick
            '//
            .Appearance.ColumnHeaderCell.TextColor = Color.DarkBlue

            'Applies back color as LightCyan for alternative records in the Grid.
            .Appearance.AlternateRecordFieldCell.BackColor = Color.LightCyan

            'Disable record preview row 
            .TableOptions.ShowRecordPreviewRow = False
            'Will enable the Group Header for the top most group.
            .TopLevelGroupOptions.ShowGroupHeader = False ' True
            'Will enable the Group Footer for the group.
            .TopLevelGroupOptions.ShowGroupFooter = False 'True
            '//
            .TableOptions.GroupHeaderSectionHeight = 30
            .TableOptions.GroupFooterSectionHeight = 30
        End With
    End Sub

    ' / -----------------------------------------------------------------
    '// Double click event for show Primary Key which hidden in Column(0)
    Private Sub GGC_TableControlCellDoubleClick(ByVal sender As Object, ByVal e As Syncfusion.Windows.Forms.Grid.Grouping.GridTableControlCellClickEventArgs) Handles GGC.TableControlCellDoubleClick
        '// Row of Column Header 
        If e.Inner.RowIndex <= 1 Then Return
        '/ Notify the double click performed in a cell.
        Dim rec As Record = Me.GGC.Table.DisplayElements(e.TableControl.CurrentCell.RowIndex).ParentRecord
        If (rec) IsNot Nothing Then MessageBox.Show("Primary key = " & rec.GetValue("CountryPK").ToString)
    End Sub

    ' / -----------------------------------------------------------------
    '// Full Select Row
    Private Sub GGC_TableControlCurrentCellActivating(ByVal sender As Object, ByVal e As Syncfusion.Windows.Forms.Grid.Grouping.GridTableControlCurrentCellActivatingEventArgs) Handles GGC.TableControlCurrentCellActivating
        '// Get Column Index 0 is the Primary Key. (Hidden column)
        e.Inner.ColIndex = 0
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    ' / -----------------------------------------------------------------
    '// Press enter each row.
    Private Sub GGC_TableControlCurrentCellKeyPress(sender As Object, e As Syncfusion.Windows.Forms.Grid.Grouping.GridTableControlKeyPressEventArgs) Handles GGC.TableControlCurrentCellKeyPress
        '// Check rows count before.
        If GGC.TableModel.RowCount <= 0 Then Return
        '/ Notify the current cell keypress 
        Dim rec As Record = Me.GGC.Table.DisplayElements(GGC.TableControl.CurrentCell.RowIndex).ParentRecord
        If (rec) IsNot Nothing Then MessageBox.Show("Primary key = " & rec.GetValue("CountryPK").ToString)
    End Sub

    Private Sub frmGGCSample_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
