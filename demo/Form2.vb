Option Explicit On 
Option Strict On

Imports DragD.QuickWhereComponent

Public Class Form2

		Inherits System.Windows.Forms.Form
		Protected QW As New QuickWhere()
		Protected APP_PATH As String

#Region " Windows Form Designer generated code "

		Public Sub New()
				MyBase.New()

				'This call is required by the Windows Form Designer.
				InitializeComponent()

				'Add any initialization after the InitializeComponent() call

		End Sub

		'Form overrides dispose to clean up the component list.
		Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
				If disposing Then
						If Not (components Is Nothing) Then
								components.Dispose()
						End If
				End If
				MyBase.Dispose(disposing)
		End Sub

		'Required by the Windows Form Designer
		Private components As System.ComponentModel.IContainer

		'NOTE: The following procedure is required by the Windows Form Designer
		'It can be modified using the Windows Form Designer.  
		'Do not modify it using the code editor.
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents txtResult As System.Windows.Forms.TextBox
	Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
	Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
	Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
	Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
	Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
	Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents optAnd As System.Windows.Forms.RadioButton
	Friend WithEvents optOr As System.Windows.Forms.RadioButton
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents btnSaveFilter As System.Windows.Forms.Button
	Friend WithEvents btnLoadFilter As System.Windows.Forms.Button
		<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container()
Me.Label2 = New System.Windows.Forms.Label()
Me.txtResult = New System.Windows.Forms.TextBox()
Me.Button1 = New System.Windows.Forms.Button()
Me.ListBox1 = New System.Windows.Forms.ListBox()
Me.ComboBox1 = New System.Windows.Forms.ComboBox()
Me.ComboBox2 = New System.Windows.Forms.ComboBox()
Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
Me.TextBox1 = New System.Windows.Forms.TextBox()
Me.TextBox2 = New System.Windows.Forms.TextBox()
Me.Panel1 = New System.Windows.Forms.Panel()
Me.optAnd = New System.Windows.Forms.RadioButton()
Me.optOr = New System.Windows.Forms.RadioButton()
Me.Label1 = New System.Windows.Forms.Label()
Me.Label3 = New System.Windows.Forms.Label()
Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
Me.btnSaveFilter = New System.Windows.Forms.Button()
Me.btnLoadFilter = New System.Windows.Forms.Button()
Me.Panel1.SuspendLayout()
Me.SuspendLayout()
'
'Label2
'
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label2.Location = New System.Drawing.Point(12, 8)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(328, 108)
Me.Label2.TabIndex = 1
Me.Label2.Text = "This example demonstrates how to:" & Microsoft.VisualBasic.ChrW(10) & "- attach various types of Win controls to Quick" & _
"Where that represent different types of values." & Microsoft.VisualBasic.ChrW(10) & "- to save the user input to a fi" & _
"le as filter for further use" & Microsoft.VisualBasic.ChrW(10) & "- to load a file as filter" & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the att" & _
"ached controls show the field name, the value type, the operation and more infor" & _
"mation."
Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtResult
'
Me.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
						Or System.Windows.Forms.AnchorStyles.Right)
Me.txtResult.Location = New System.Drawing.Point(12, 416)
Me.txtResult.Multiline = True
Me.txtResult.Name = "txtResult"
Me.txtResult.Size = New System.Drawing.Size(328, 68)
Me.txtResult.TabIndex = 4
Me.txtResult.Text = ""
'
'Button1
'
Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Button1.Location = New System.Drawing.Point(12, 376)
Me.Button1.Name = "Button1"
Me.Button1.Size = New System.Drawing.Size(152, 23)
Me.Button1.TabIndex = 5
Me.Button1.Text = "Generate WHERE clause"
'
'ListBox1
'
Me.ListBox1.Items.AddRange(New Object() {"ListValue1", "ListValue2", "ListValue3", "ListValue4", "ListValue5", "ListValue6"})
Me.ListBox1.Location = New System.Drawing.Point(12, 132)
Me.ListBox1.Name = "ListBox1"
Me.ListBox1.Size = New System.Drawing.Size(120, 82)
Me.ListBox1.TabIndex = 6
Me.ToolTip1.SetToolTip(Me.ListBox1, "Field1, Contains, String. Select an item to include it into the WHERE clause")
'
'ComboBox1
'
Me.ComboBox1.Items.AddRange(New Object() {"", "10", "20", "30", "40", "50"})
Me.ComboBox1.Location = New System.Drawing.Point(12, 228)
Me.ComboBox1.Name = "ComboBox1"
Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
Me.ComboBox1.TabIndex = 7
Me.ToolTip1.SetToolTip(Me.ComboBox1, "Field2, Greater or  Equal, Numeric. Select an item to include it into the WHERE c" & _
"lause")
'
'ComboBox2
'
Me.ComboBox2.Items.AddRange(New Object() {"Yes", "No"})
Me.ComboBox2.Location = New System.Drawing.Point(12, 264)
Me.ComboBox2.Name = "ComboBox2"
Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
Me.ComboBox2.TabIndex = 8
Me.ToolTip1.SetToolTip(Me.ComboBox2, "Field3, Equal, Boolean. Select an item to include it into the WHERE clause")
'
'MonthCalendar1
'
Me.MonthCalendar1.Location = New System.Drawing.Point(146, 132)
Me.MonthCalendar1.Name = "MonthCalendar1"
Me.MonthCalendar1.TabIndex = 9
Me.ToolTip1.SetToolTip(Me.MonthCalendar1, "Field4, Between, Date. Select a single date or a period")
'
'TextBox1
'
Me.TextBox1.Location = New System.Drawing.Point(56, 296)
Me.TextBox1.Name = "TextBox1"
Me.TextBox1.Size = New System.Drawing.Size(47, 20)
Me.TextBox1.TabIndex = 10
Me.TextBox1.Text = "15.09"
Me.ToolTip1.SetToolTip(Me.TextBox1, "Field5, Not Between, Numeric. Change the numeric values or leave them blank")
'
'TextBox2
'
Me.TextBox2.Location = New System.Drawing.Point(56, 328)
Me.TextBox2.Name = "TextBox2"
Me.TextBox2.Size = New System.Drawing.Size(47, 20)
Me.TextBox2.TabIndex = 11
Me.TextBox2.Text = "23.98"
Me.ToolTip1.SetToolTip(Me.TextBox2, "Field5, Not Between, Numeric. Change the numeric values or leave them blank")
'
'Panel1
'
Me.Panel1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.optAnd, Me.optOr})
Me.Panel1.Location = New System.Drawing.Point(206, 371)
Me.Panel1.Name = "Panel1"
Me.Panel1.Size = New System.Drawing.Size(132, 28)
Me.Panel1.TabIndex = 12
Me.ToolTip1.SetToolTip(Me.Panel1, "Select the Operator of the WHERE clause")
'
'optAnd
'
Me.optAnd.Checked = True
Me.optAnd.Location = New System.Drawing.Point(12, 7)
Me.optAnd.Name = "optAnd"
Me.optAnd.Size = New System.Drawing.Size(48, 16)
Me.optAnd.TabIndex = 0
Me.optAnd.TabStop = True
Me.optAnd.Text = "AND"
Me.ToolTip1.SetToolTip(Me.optAnd, "Select the Operator of the WHERE clause")
'
'optOr
'
Me.optOr.Location = New System.Drawing.Point(76, 7)
Me.optOr.Name = "optOr"
Me.optOr.Size = New System.Drawing.Size(44, 16)
Me.optOr.TabIndex = 0
Me.optOr.Text = "OR"
Me.ToolTip1.SetToolTip(Me.optOr, "Select the Operator of the WHERE clause")
'
'Label1
'
Me.Label1.Location = New System.Drawing.Point(12, 300)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(40, 16)
Me.Label1.TabIndex = 13
Me.Label1.Text = "From:"
'
'Label3
'
Me.Label3.Location = New System.Drawing.Point(12, 328)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(32, 16)
Me.Label3.TabIndex = 13
Me.Label3.Text = "To:"
'
'ToolTip1
'
Me.ToolTip1.AutoPopDelay = 10000
Me.ToolTip1.InitialDelay = 500
Me.ToolTip1.ReshowDelay = 100
Me.ToolTip1.ShowAlways = True
'
'btnSaveFilter
'
Me.btnSaveFilter.Location = New System.Drawing.Point(246, 296)
Me.btnSaveFilter.Name = "btnSaveFilter"
Me.btnSaveFilter.Size = New System.Drawing.Size(92, 23)
Me.btnSaveFilter.TabIndex = 14
Me.btnSaveFilter.Text = "Save Filter"
Me.ToolTip1.SetToolTip(Me.btnSaveFilter, "Save to {AppPath}\QuickWhereFilter.xml")
'
'btnLoadFilter
'
Me.btnLoadFilter.Location = New System.Drawing.Point(246, 324)
Me.btnLoadFilter.Name = "btnLoadFilter"
Me.btnLoadFilter.Size = New System.Drawing.Size(92, 23)
Me.btnLoadFilter.TabIndex = 15
Me.btnLoadFilter.Text = "Load Filter"
Me.ToolTip1.SetToolTip(Me.btnLoadFilter, "Load from {AppPath}\QuickWhereFilter.xml")
'
'Form2
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(352, 493)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnLoadFilter, Me.btnSaveFilter, Me.Label1, Me.Panel1, Me.TextBox2, Me.TextBox1, Me.MonthCalendar1, Me.ComboBox2, Me.ComboBox1, Me.ListBox1, Me.Button1, Me.txtResult, Me.Label2, Me.Label3})
Me.Name = "Form2"
Me.Text = "QuickWhere component demo, Example 2"
Me.Panel1.ResumeLayout(False)
Me.ResumeLayout(False)

		End Sub

#End Region

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		APP_PATH = System.IO.Path.GetDirectoryName( _
								System.IO.Path.GetDirectoryName( _
								Application.ExecutablePath))

		With QW
			.Add("Field1", WhereConditions.Contains, TypeOfValues.StringType, ListBox1)
			.Add("Field2", WhereConditions.GreaterOrEqual, TypeOfValues.NumericType, ComboBox1)
			.Add("Field3", TypeOfValues.BooleanType, ComboBox2)
			.Add("Field4", "TableName1", TypeOfValues.DateType, WhereConditions.Between, MonthCalendar1, MonthCalendar1)
			.Add("Field5", "TableName2", TypeOfValues.NumericType, WhereConditions.NotBetween, TextBox1, TextBox2)
		End With

	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
	Dim anOperator As Operators = Operators.AndOperator

		If optOr.Checked Then anOperator = Operators.OrOperator
		txtResult.Text = QW.GetSql(anOperator)

	End Sub

	Private Sub btnSaveFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFilter.Click
		Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

		QuickWhere.SaveFilter(APP_PATH & "\QuickWhereFilter.xml", QW)

		Cursor.Current = System.Windows.Forms.Cursors.Default

	End Sub

Private Sub btnLoadFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadFilter.Click
		Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

		' it is necessary to set the form which controls belong to
		QW.ParentForm = Me
		QW.LoadFilter(APP_PATH & "\QuickWhereFilter.xml", Me)

		Cursor.Current = System.Windows.Forms.Cursors.Default

	End Sub

	Private Sub Form2_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		QW = Nothing
	End Sub

Private Function TryMe() As WhereClause
	Dim wiFirstName As New WhereItem()
	Dim wiLastName As New WhereItem()
	Dim wiDepartment As New WhereItem()
	Dim wiSale As New WhereItem()
	Dim wiSaleDate As New WhereItem()
	Dim wcSearchByName As New WhereClause()
	Dim wcSearchBySale As New WhereClause()
	Dim wcFinalSearch As New WhereClause()

	' The default WhereCondition is WhereConditions.EqualTo
	' The default ValueType is TypeOfValues.StringType

	With wiFirstName
		.FieldName = "FirstName"
		.TableName = "Personnel"
		.Value = "John"
	End With

	With wiLastName
		.FieldName = "LastName"
		.TableName = "Personnel"
		.Value = "son"
		.WhereCondition = WhereConditions.EndsWith
	End With

	With wiDepartment
		.FieldName = "DepartmentID"
		.TableName = "Departments"
		.Value = "9"
		.ValueType = TypeOfValues.NumericType
	End With

	' The default Operators is Operators.AndOperator
	wcSearchByName.Add(wiFirstName)
	wcSearchByName.Add(wiLastName)
	wcSearchByName.Add(wiDepartment)


	With wiSale
		.FieldName = "Sale"
		.Value = "15000"
		.TableName = "Sales"
		.ValueType = TypeOfValues.NumericType
		.WhereCondition = WhereConditions.GreaterOrEqual
	End With

	With wiSaleDate
		.FieldName = "SaleDate"
		.TableName = "Sales"
		.Value = "01/10/87"
		.Value2 = "25/10/87"
		.ValueType = TypeOfValues.DateType
		.WhereCondition = WhereConditions.Between
	End With
	wcSearchBySale.Add(wiSale)
	wcSearchBySale.Add(wiSaleDate)


	wcFinalSearch.Add(wcSearchByName)
	wcFinalSearch.Add(wcSearchBySale)

	MsgBox(wcFinalSearch.GetSql(Operators.OrOperator))

	Return wcFinalSearch

End Function

End Class
