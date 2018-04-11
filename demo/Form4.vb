Option Explicit On 
Option Strict On

Imports DragD.QuickWhereComponent

Public Class Form4

		Inherits System.Windows.Forms.Form
		Protected QW As QuickWhere

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
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
	Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
	Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
	Friend WithEvents txtDateMarker As System.Windows.Forms.TextBox
	Friend WithEvents txtStringMarker As System.Windows.Forms.TextBox
	Friend WithEvents txtStringDelimiter As System.Windows.Forms.TextBox
	Friend WithEvents txtWildCardMarker As System.Windows.Forms.TextBox
		<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container()
Me.Label0 = New System.Windows.Forms.Label()
Me.txtResult = New System.Windows.Forms.TextBox()
Me.Button1 = New System.Windows.Forms.Button()
Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
Me.ListBox1 = New System.Windows.Forms.ListBox()
Me.Panel1 = New System.Windows.Forms.Panel()
Me.Label1 = New System.Windows.Forms.Label()
Me.Label2 = New System.Windows.Forms.Label()
Me.txtDateMarker = New System.Windows.Forms.TextBox()
Me.Label3 = New System.Windows.Forms.Label()
Me.txtStringMarker = New System.Windows.Forms.TextBox()
Me.txtStringDelimiter = New System.Windows.Forms.TextBox()
Me.Label4 = New System.Windows.Forms.Label()
Me.txtWildCardMarker = New System.Windows.Forms.TextBox()
Me.Label5 = New System.Windows.Forms.Label()
Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
Me.TextBox1 = New System.Windows.Forms.TextBox()
Me.Panel1.SuspendLayout()
Me.SuspendLayout()
'
'Label0
'
Me.Label0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label0.Location = New System.Drawing.Point(12, 12)
Me.Label0.Name = "Label0"
Me.Label0.Size = New System.Drawing.Size(284, 72)
Me.Label0.TabIndex = 1
Me.Label0.Text = "This example demonstrates how the general setting of QuickWhere reflect the gener" & _
"ated WHERE clause." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the controls show the field name, the value " & _
"type, the operation and more information."
'
'txtResult
'
Me.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
						Or System.Windows.Forms.AnchorStyles.Right)
Me.txtResult.Location = New System.Drawing.Point(12, 304)
Me.txtResult.Multiline = True
Me.txtResult.Name = "txtResult"
Me.txtResult.Size = New System.Drawing.Size(288, 68)
Me.txtResult.TabIndex = 4
Me.txtResult.Text = ""
'
'Button1
'
Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Button1.Location = New System.Drawing.Point(12, 264)
Me.Button1.Name = "Button1"
Me.Button1.Size = New System.Drawing.Size(148, 23)
Me.Button1.TabIndex = 5
Me.Button1.Text = "Generate WHERE clause"
'
'ListBox1
'
Me.ListBox1.Items.AddRange(New Object() {"ListValue1", "ListValue2", "ListValue3", "ListValue4", "ListValue5", "ListValue6"})
Me.ListBox1.Location = New System.Drawing.Point(12, 157)
Me.ListBox1.Name = "ListBox1"
Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
Me.ListBox1.Size = New System.Drawing.Size(84, 82)
Me.ListBox1.TabIndex = 25
Me.ToolTip1.SetToolTip(Me.ListBox1, "Field3, IN (...), String. Select an item to include it into the WHERE clause")
'
'Panel1
'
Me.Panel1.BackColor = System.Drawing.Color.Silver
Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.Label2, Me.txtDateMarker, Me.Label3, Me.txtStringMarker, Me.txtStringDelimiter, Me.Label4, Me.txtWildCardMarker, Me.Label5})
Me.Panel1.Location = New System.Drawing.Point(140, 93)
Me.Panel1.Name = "Panel1"
Me.Panel1.Size = New System.Drawing.Size(148, 148)
Me.Panel1.TabIndex = 7
Me.ToolTip1.SetToolTip(Me.Panel1, "Change the general settings and regenerate the WHERE clause")
'
'Label1
'
Me.Label1.BackColor = System.Drawing.Color.Silver
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label1.Location = New System.Drawing.Point(22, 6)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(108, 28)
Me.Label1.TabIndex = 0
Me.Label1.Text = "QuickWhere General Settings"
Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.ToolTip1.SetToolTip(Me.Label1, "Change the general settings and regenerate the WHERE clause")
'
'Label2
'
Me.Label2.BackColor = System.Drawing.Color.Silver
Me.Label2.Location = New System.Drawing.Point(12, 43)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(100, 20)
Me.Label2.TabIndex = 2
Me.Label2.Text = "Date Marker"
Me.ToolTip1.SetToolTip(Me.Label2, "Change the general settings and regenerate the WHERE clause")
'
'txtDateMarker
'
Me.txtDateMarker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtDateMarker.Location = New System.Drawing.Point(114, 43)
Me.txtDateMarker.Name = "txtDateMarker"
Me.txtDateMarker.Size = New System.Drawing.Size(24, 20)
Me.txtDateMarker.TabIndex = 1
Me.txtDateMarker.Text = "#"
Me.txtDateMarker.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.ToolTip1.SetToolTip(Me.txtDateMarker, "Change the general settings and regenerate the WHERE clause")
'
'Label3
'
Me.Label3.BackColor = System.Drawing.Color.Silver
Me.Label3.Location = New System.Drawing.Point(12, 67)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(100, 20)
Me.Label3.TabIndex = 2
Me.Label3.Text = "String Marker"
Me.ToolTip1.SetToolTip(Me.Label3, "Change the general settings and regenerate the WHERE clause")
'
'txtStringMarker
'
Me.txtStringMarker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtStringMarker.Location = New System.Drawing.Point(114, 67)
Me.txtStringMarker.Name = "txtStringMarker"
Me.txtStringMarker.Size = New System.Drawing.Size(24, 20)
Me.txtStringMarker.TabIndex = 1
Me.txtStringMarker.Text = """"
Me.txtStringMarker.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.ToolTip1.SetToolTip(Me.txtStringMarker, "Change the general settings and regenerate the WHERE clause")
'
'txtStringDelimiter
'
Me.txtStringDelimiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtStringDelimiter.Location = New System.Drawing.Point(114, 91)
Me.txtStringDelimiter.Name = "txtStringDelimiter"
Me.txtStringDelimiter.Size = New System.Drawing.Size(24, 20)
Me.txtStringDelimiter.TabIndex = 1
Me.txtStringDelimiter.Text = ","
Me.txtStringDelimiter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.ToolTip1.SetToolTip(Me.txtStringDelimiter, "Change the general settings and regenerate the WHERE clause")
'
'Label4
'
Me.Label4.BackColor = System.Drawing.Color.Silver
Me.Label4.Location = New System.Drawing.Point(12, 91)
Me.Label4.Name = "Label4"
Me.Label4.Size = New System.Drawing.Size(100, 20)
Me.Label4.TabIndex = 2
Me.Label4.Text = "String Delimiter"
Me.ToolTip1.SetToolTip(Me.Label4, "Change the general settings and regenerate the WHERE clause")
'
'txtWildCardMarker
'
Me.txtWildCardMarker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtWildCardMarker.Location = New System.Drawing.Point(114, 115)
Me.txtWildCardMarker.Name = "txtWildCardMarker"
Me.txtWildCardMarker.Size = New System.Drawing.Size(24, 20)
Me.txtWildCardMarker.TabIndex = 1
Me.txtWildCardMarker.Text = "*"
Me.txtWildCardMarker.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
Me.ToolTip1.SetToolTip(Me.txtWildCardMarker, "Change the general settings and regenerate the WHERE clause")
'
'Label5
'
Me.Label5.BackColor = System.Drawing.Color.Silver
Me.Label5.Location = New System.Drawing.Point(12, 115)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(100, 20)
Me.Label5.TabIndex = 2
Me.Label5.Text = "Wild Card Marker"
Me.ToolTip1.SetToolTip(Me.Label5, "Change the general settings and regenerate the WHERE clause")
'
'DateTimePicker1
'
Me.DateTimePicker1.CustomFormat = ""
Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short
Me.DateTimePicker1.Location = New System.Drawing.Point(12, 125)
Me.DateTimePicker1.Name = "DateTimePicker1"
Me.DateTimePicker1.Size = New System.Drawing.Size(88, 20)
Me.DateTimePicker1.TabIndex = 9
Me.ToolTip1.SetToolTip(Me.DateTimePicker1, "Field2, Equal, Date")
Me.DateTimePicker1.Value = New Date(1987, 5, 22, 0, 0, 0, 0)
'
'TextBox1
'
Me.TextBox1.Location = New System.Drawing.Point(12, 93)
Me.TextBox1.Name = "TextBox1"
Me.TextBox1.Size = New System.Drawing.Size(88, 20)
Me.TextBox1.TabIndex = 8
Me.TextBox1.Text = "someStringValue"
Me.ToolTip1.SetToolTip(Me.TextBox1, "Field1, Like, String")
'
'Form4
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(312, 381)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListBox1, Me.DateTimePicker1, Me.TextBox1, Me.Panel1, Me.Button1, Me.txtResult, Me.Label0})
Me.Name = "Form4"
Me.Text = "QuickWhere component demo, Example 4"
Me.Panel1.ResumeLayout(False)
Me.ResumeLayout(False)

		End Sub

#End Region

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
	Dim wcListBox As New WhereControl()

		QW = New QuickWhere()

		With wcListBox
			.FieldName = "Field3"
			.WhereCondition = WhereConditions.InCondition
			.Value = GetStringFromListBox(ListBox1)
		End With

		With QW
			.SetGenerals(CChar(txtStringMarker.Text), _
									CChar(txtDateMarker.Text), , _
									CChar(txtWildCardMarker.Text), _
									CChar(txtStringDelimiter.Text))

			.Add("Field1", "Table1", TypeOfValues.StringType, WhereConditions.BeginsWith, TextBox1)
			.Add("Field2", TypeOfValues.DateType, DateTimePicker1)
			.Add(wcListBox)
		End With

		txtResult.Text = QW.GetSql

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

	Private Sub Form4_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		QW = Nothing
	End Sub
End Class
