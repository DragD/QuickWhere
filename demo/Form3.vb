Option Explicit On 
Option Strict On

Imports DragD.QuickWhereComponent

Public Class Form3

		Inherits System.Windows.Forms.Form

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
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents txtResult As System.Windows.Forms.TextBox
	Friend WithEvents Label0 As System.Windows.Forms.Label
	Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents Panel4 As System.Windows.Forms.Panel
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
	Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents chkCountable As System.Windows.Forms.CheckBox
	Friend WithEvents chkReadable As System.Windows.Forms.CheckBox
	Friend WithEvents chkApplicable As System.Windows.Forms.CheckBox
	Friend WithEvents chkHasValues As System.Windows.Forms.CheckBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Panel3 As System.Windows.Forms.Panel
	Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents Label7 As System.Windows.Forms.Label
	Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
	Friend WithEvents Button2 As System.Windows.Forms.Button
	Friend WithEvents chkUsed1 As System.Windows.Forms.CheckBox
	Friend WithEvents chkUsed2 As System.Windows.Forms.CheckBox
	Friend WithEvents chkUsed3 As System.Windows.Forms.CheckBox
		<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container()
Me.Label0 = New System.Windows.Forms.Label()
Me.txtResult = New System.Windows.Forms.TextBox()
Me.Button1 = New System.Windows.Forms.Button()
Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
Me.TextBox1 = New System.Windows.Forms.TextBox()
Me.ListBox1 = New System.Windows.Forms.ListBox()
Me.chkCountable = New System.Windows.Forms.CheckBox()
Me.chkReadable = New System.Windows.Forms.CheckBox()
Me.chkApplicable = New System.Windows.Forms.CheckBox()
Me.chkHasValues = New System.Windows.Forms.CheckBox()
Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
Me.TextBox2 = New System.Windows.Forms.TextBox()
Me.chkUsed1 = New System.Windows.Forms.CheckBox()
Me.chkUsed2 = New System.Windows.Forms.CheckBox()
Me.chkUsed3 = New System.Windows.Forms.CheckBox()
Me.Panel4 = New System.Windows.Forms.Panel()
Me.Label7 = New System.Windows.Forms.Label()
Me.Panel1 = New System.Windows.Forms.Panel()
Me.Label2 = New System.Windows.Forms.Label()
Me.Panel2 = New System.Windows.Forms.Panel()
Me.Label1 = New System.Windows.Forms.Label()
Me.Panel3 = New System.Windows.Forms.Panel()
Me.Label3 = New System.Windows.Forms.Label()
Me.Button2 = New System.Windows.Forms.Button()
Me.Panel4.SuspendLayout()
Me.Panel1.SuspendLayout()
Me.Panel2.SuspendLayout()
Me.Panel3.SuspendLayout()
Me.SuspendLayout()
'
'Label0
'
Me.Label0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label0.Location = New System.Drawing.Point(12, 12)
Me.Label0.Name = "Label0"
Me.Label0.Size = New System.Drawing.Size(416, 120)
Me.Label0.TabIndex = 1
Me.Label0.Text = "This example demonstrates:" & Microsoft.VisualBasic.ChrW(10) & " - the use of IN(...) clause generated from ListBox or" & _
" CheckedListBox" & Microsoft.VisualBasic.ChrW(10) & " - how to use mixed ANDs and ORs in two alternative ways" & Microsoft.VisualBasic.ChrW(10) & " - how " & _
"to use operation over values" & Microsoft.VisualBasic.ChrW(10) & " - how to use IS (NOT) NULL clause" & Microsoft.VisualBasic.ChrW(10) & " - the applicati" & _
"on of Used property" & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the controls show the field name, the value" & _
" type, the operation and more information."
'
'txtResult
'
Me.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
						Or System.Windows.Forms.AnchorStyles.Right)
Me.txtResult.Location = New System.Drawing.Point(12, 416)
Me.txtResult.Multiline = True
Me.txtResult.Name = "txtResult"
Me.txtResult.Size = New System.Drawing.Size(416, 68)
Me.txtResult.TabIndex = 4
Me.txtResult.Text = ""
'
'Button1
'
Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Button1.Location = New System.Drawing.Point(12, 376)
Me.Button1.Name = "Button1"
Me.Button1.Size = New System.Drawing.Size(200, 23)
Me.Button1.TabIndex = 5
Me.Button1.Text = "Generate WHERE clause"
'
'TextBox1
'
Me.TextBox1.Location = New System.Drawing.Point(7, 114)
Me.TextBox1.Name = "TextBox1"
Me.TextBox1.Size = New System.Drawing.Size(84, 20)
Me.TextBox1.TabIndex = 25
Me.TextBox1.Text = "someTextValue"
Me.ToolTip1.SetToolTip(Me.TextBox1, "Field2, Equal, String, Mid operation over the value. Change the text in the textb" & _
"ox")
'
'ListBox1
'
Me.ListBox1.Items.AddRange(New Object() {"ListValue1", "ListValue2", "ListValue3", "ListValue4", "ListValue5", "ListValue6"})
Me.ListBox1.Location = New System.Drawing.Point(7, 24)
Me.ListBox1.Name = "ListBox1"
Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
Me.ListBox1.Size = New System.Drawing.Size(84, 82)
Me.ListBox1.TabIndex = 24
Me.ToolTip1.SetToolTip(Me.ListBox1, "Field1, IN (...), String. Select an item to include it into the WHERE clause")
'
'chkCountable
'
Me.chkCountable.Location = New System.Drawing.Point(9, 80)
Me.chkCountable.Name = "chkCountable"
Me.chkCountable.Size = New System.Drawing.Size(87, 24)
Me.chkCountable.TabIndex = 35
Me.chkCountable.Text = "Countable"
Me.ToolTip1.SetToolTip(Me.chkCountable, "Field5_Countable, Equal, Boolean")
'
'chkReadable
'
Me.chkReadable.Location = New System.Drawing.Point(9, 108)
Me.chkReadable.Name = "chkReadable"
Me.chkReadable.Size = New System.Drawing.Size(87, 24)
Me.chkReadable.TabIndex = 33
Me.chkReadable.Text = "Readable"
Me.ToolTip1.SetToolTip(Me.chkReadable, "Field5_Readable, Equal, Boolean")
'
'chkApplicable
'
Me.chkApplicable.Location = New System.Drawing.Point(9, 52)
Me.chkApplicable.Name = "chkApplicable"
Me.chkApplicable.Size = New System.Drawing.Size(87, 24)
Me.chkApplicable.TabIndex = 32
Me.chkApplicable.Text = "Applicable"
Me.ToolTip1.SetToolTip(Me.chkApplicable, "Field4_Applicable, Equal, Boolean")
'
'chkHasValues
'
Me.chkHasValues.Location = New System.Drawing.Point(9, 24)
Me.chkHasValues.Name = "chkHasValues"
Me.chkHasValues.Size = New System.Drawing.Size(87, 24)
Me.chkHasValues.TabIndex = 34
Me.chkHasValues.Text = "Has Values"
Me.ToolTip1.SetToolTip(Me.chkHasValues, "Field3, (NOT) IS NULL,Value type doesn't matter")
'
'CheckedListBox1
'
Me.CheckedListBox1.CheckOnClick = True
Me.CheckedListBox1.Items.AddRange(New Object() {"1001", "1002", "1003", "1004", "1005"})
Me.CheckedListBox1.Location = New System.Drawing.Point(17, 24)
Me.CheckedListBox1.Name = "CheckedListBox1"
Me.CheckedListBox1.Size = New System.Drawing.Size(68, 79)
Me.CheckedListBox1.TabIndex = 34
Me.ToolTip1.SetToolTip(Me.CheckedListBox1, "Field7, IN (...), String. Check an item to include it into the WHERE clause")
'
'TextBox2
'
Me.TextBox2.Location = New System.Drawing.Point(10, 114)
Me.TextBox2.Name = "TextBox2"
Me.TextBox2.Size = New System.Drawing.Size(86, 20)
Me.TextBox2.TabIndex = 33
Me.TextBox2.Text = "22/05/87"
Me.ToolTip1.SetToolTip(Me.TextBox2, "Field8, Equal, Date, Format operation over the value. Change the text in the text" & _
"box")
'
'chkUsed1
'
Me.chkUsed1.Checked = True
Me.chkUsed1.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkUsed1.Location = New System.Drawing.Point(38, 188)
Me.chkUsed1.Name = "chkUsed1"
Me.chkUsed1.Size = New System.Drawing.Size(56, 16)
Me.chkUsed1.TabIndex = 36
Me.chkUsed1.Text = "Used"
Me.ToolTip1.SetToolTip(Me.chkUsed1, "Turns ON/OFF the use of the controls on the panel. Applies for the alternative wa" & _
"y only")
'
'chkUsed2
'
Me.chkUsed2.Checked = True
Me.chkUsed2.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkUsed2.Location = New System.Drawing.Point(168, 188)
Me.chkUsed2.Name = "chkUsed2"
Me.chkUsed2.Size = New System.Drawing.Size(56, 16)
Me.chkUsed2.TabIndex = 36
Me.chkUsed2.Text = "Used"
Me.ToolTip1.SetToolTip(Me.chkUsed2, "Turns ON/OFF the use of the controls on the panel. Applies for the alternative wa" & _
"y only")
'
'chkUsed3
'
Me.chkUsed3.Checked = True
Me.chkUsed3.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkUsed3.Location = New System.Drawing.Point(308, 188)
Me.chkUsed3.Name = "chkUsed3"
Me.chkUsed3.Size = New System.Drawing.Size(56, 16)
Me.chkUsed3.TabIndex = 36
Me.chkUsed3.Text = "Used"
Me.ToolTip1.SetToolTip(Me.chkUsed3, "Turns ON/OFF the use of the controls on the panel. Applies for the alternative wa" & _
"y only")
'
'Panel4
'
Me.Panel4.BackColor = System.Drawing.Color.Gray
Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.Panel4.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkUsed1, Me.Label7, Me.Panel1, Me.Panel2, Me.Panel3, Me.chkUsed2, Me.chkUsed3})
Me.Panel4.Location = New System.Drawing.Point(12, 143)
Me.Panel4.Name = "Panel4"
Me.Panel4.Size = New System.Drawing.Size(416, 220)
Me.Panel4.TabIndex = 32
'
'Label7
'
Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label7.Location = New System.Drawing.Point(16, 8)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(372, 14)
Me.Label7.TabIndex = 35
Me.Label7.Text = "OR"
Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
'
'Panel1
'
Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.TextBox1, Me.ListBox1})
Me.Panel1.Location = New System.Drawing.Point(16, 32)
Me.Panel1.Name = "Panel1"
Me.Panel1.Size = New System.Drawing.Size(104, 144)
Me.Panel1.TabIndex = 34
'
'Label2
'
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label2.Location = New System.Drawing.Point(32, 7)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(32, 12)
Me.Label2.TabIndex = 31
Me.Label2.Text = "AND"
'
'Panel2
'
Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkCountable, Me.chkReadable, Me.chkApplicable, Me.chkHasValues, Me.Label1})
Me.Panel2.Location = New System.Drawing.Point(146, 32)
Me.Panel2.Name = "Panel2"
Me.Panel2.Size = New System.Drawing.Size(108, 144)
Me.Panel2.TabIndex = 33
'
'Label1
'
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label1.Location = New System.Drawing.Point(36, 7)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(32, 12)
Me.Label1.TabIndex = 31
Me.Label1.Text = "AND"
'
'Panel3
'
Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.CheckedListBox1, Me.TextBox2, Me.Label3})
Me.Panel3.Location = New System.Drawing.Point(280, 32)
Me.Panel3.Name = "Panel3"
Me.Panel3.Size = New System.Drawing.Size(108, 144)
Me.Panel3.TabIndex = 32
'
'Label3
'
Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label3.Location = New System.Drawing.Point(36, 7)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(32, 12)
Me.Label3.TabIndex = 31
Me.Label3.Text = "AND"
'
'Button2
'
Me.Button2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Button2.Location = New System.Drawing.Point(228, 376)
Me.Button2.Name = "Button2"
Me.Button2.Size = New System.Drawing.Size(200, 23)
Me.Button2.TabIndex = 33
Me.Button2.Text = "Generate WHERE - alternative way"
'
'Form3
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(440, 493)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button2, Me.Panel4, Me.Button1, Me.txtResult, Me.Label0})
Me.Name = "Form3"
Me.Text = "QuickWhere component demo, Example 3"
Me.Panel4.ResumeLayout(False)
Me.Panel1.ResumeLayout(False)
Me.Panel2.ResumeLayout(False)
Me.Panel3.ResumeLayout(False)
Me.ResumeLayout(False)

		End Sub

#End Region

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
	Dim QW1 As New QuickWhere()
	Dim QW2 As New QuickWhere()
	Dim QW3 As New QuickWhere()

	Dim wcListBox As New WhereControl()
	Dim wcCheckedListBox As New WhereControl()

	Dim indValueOperation As Integer

		With wcListBox
			.FieldName = "Field1"
			.WhereCondition = WhereConditions.InCondition
			.Value = GetStringFromListBox(ListBox1)
		End With
		QW1.Add(wcListBox)


		indValueOperation = QW1.Add("Field2", TextBox1)
		QW1(indValueOperation).ValueOperation = "Mid(@@,2,3)"

		With QW2
			Dim wcHasValues As New WhereControl()
			wcHasValues.FieldName = "Field3"

			If chkHasValues.Checked Then
				wcHasValues.WhereCondition = WhereConditions.NotIsNull
			Else
				wcHasValues.WhereCondition = WhereConditions.IsNull
			End If

			.Add(wcHasValues)
			.Add("Field4_Applicable", TypeOfValues.BooleanType, chkApplicable)
			.Add("Field5_Countable", TypeOfValues.BooleanType, chkCountable)
			.Add("Field6_Readable", TypeOfValues.BooleanType, chkReadable)

		End With


		With wcCheckedListBox
			.FieldName = "Field7"
			.WhereCondition = WhereConditions.InCondition
			.ValueType = TypeOfValues.NumericType
			.Value = GetStringFromCheckedListBox(CheckedListBox1)
		End With
		QW3.Add(wcCheckedListBox)

		indValueOperation = QW3.Add("Field8", TypeOfValues.DateType, TextBox2)
		QW3(indValueOperation).ValueOperation = "Format(@@,LongDate)"

		txtResult.Text = "(" & QW1.GetSql & ") OR (" & QW2.GetSql & ") OR (" & QW3.GetSql & ")"

		QW1 = Nothing
		QW2 = Nothing
		QW3 = Nothing
		wcListBox = Nothing
		wcCheckedListBox = Nothing

	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
	Dim wi As New WhereItem()
	Dim WhereClause1 As New WhereClause(Operators.AndOperator)
	Dim WhereClause2 As New WhereClause(Operators.AndOperator)
	Dim WhereClause3 As New WhereClause(Operators.AndOperator)
	Dim WhereClauseFinal As New WhereClause(Operators.OrOperator)
	Dim qw As New QuickWhere()

		Call QuickWhere.SetGenerals()

		With wi
			.FieldName = "Field1"
			.WhereCondition = WhereConditions.InCondition
			.Value = GetStringFromListBox(ListBox1)
		End With
		WhereClause1.Add(wi.Copy) ' use Copy method to create identical copy of the WhereItem

		wi = New WhereItem()
		With wi
			 .FieldName = "Field2"
			 .Value = TextBox1.Text
			 .ValueOperation = "Mid(@@,2,3)"
		End With
		WhereClause1.Add(wi.Copy)


		wi = New WhereItem()
		With wi
			 .FieldName = "Field3"
			If chkHasValues.Checked Then
				.WhereCondition = WhereConditions.NotIsNull
			Else
				.WhereCondition = WhereConditions.IsNull
			End If
		End With
		WhereClause2.Add(wi.Copy)

		With qw
			.Add("Field4_Applicable", TypeOfValues.BooleanType, chkApplicable)
			WhereClause2.Add(.Item(0))

			.Add("Field5_Countable", TypeOfValues.BooleanType, chkCountable)
			WhereClause2.Add(.Item(1))

			.Add("Field6_Readable", TypeOfValues.BooleanType, chkReadable)
			WhereClause2.Add(.Item(2))
		End With

		wi = New WhereItem()
		With wi
			.FieldName = "Field7"
			.WhereCondition = WhereConditions.InCondition
			.ValueType = TypeOfValues.NumericType
			.Value = GetStringFromCheckedListBox(CheckedListBox1)
		End With
		WhereClause3.Add(wi.Copy)

		wi = New WhereItem()
		With wi
			.FieldName = "Field8"
			.ValueType = TypeOfValues.DateType
			.Value = TextBox2.Text
			.ValueOperation = "Format(@@,LongDate)"
		End With
		WhereClause3.Add(wi.Copy)

		WhereClause1.Used = chkUsed1.Checked
		WhereClause2.Used = chkUsed2.Checked
		WhereClause3.Used = chkUsed3.Checked

		WhereClauseFinal.Add(WhereClause1)
		WhereClauseFinal.Add(WhereClause2)
		WhereClauseFinal.Add(WhereClause3)
		txtResult.Text = WhereClauseFinal.GetSql

		wi = Nothing
		WhereClause1 = Nothing
		WhereClause2 = Nothing
		WhereClause3 = Nothing
		WhereClauseFinal = Nothing
		qw = Nothing

	End Sub

	Private Function GetStringFromListBox(ByVal lst As ListBox) As String
	Dim enumList As IEnumerator
	Dim sBuilder As New System.Text.StringBuilder()

			enumList = lst.SelectedItems.GetEnumerator
			While enumList.MoveNext
				sBuilder.Append("," & Convert.ToString(enumList.Current))
			End While

			Return Mid(sBuilder.ToString, 2)

		sBuilder = Nothing

	End Function

	Private Function GetStringFromCheckedListBox(ByVal chb As CheckedListBox) As String
	Dim sBuilder As New System.Text.StringBuilder()
	Dim I As Integer
	Dim CheckedCount As Integer = chb.CheckedItems.Count

				For I = 0 To CheckedCount - 1
					sBuilder.Append("," & Convert.ToString(chb.CheckedItems(I)))
				Next

				Return Mid(sBuilder.ToString, 2)

		sBuilder = Nothing

	End Function

End Class
