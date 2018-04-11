Option Explicit On 
Option Strict On

Imports DragD.QuickWhereComponent

Public Class Form1

		Inherits System.Windows.Forms.Form
		Protected QW As New QuickWhere()

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
	Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
	Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents txtResult As System.Windows.Forms.TextBox
		<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.Label2 = New System.Windows.Forms.Label()
Me.TextBox1 = New System.Windows.Forms.TextBox()
Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
Me.txtResult = New System.Windows.Forms.TextBox()
Me.Button1 = New System.Windows.Forms.Button()
Me.SuspendLayout()
'
'Label2
'
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label2.Location = New System.Drawing.Point(12, 12)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(292, 36)
Me.Label2.TabIndex = 1
Me.Label2.Text = "This is very simple example. It demonstrates the basic use of QuickWhere."
Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'TextBox1
'
Me.TextBox1.Location = New System.Drawing.Point(12, 60)
Me.TextBox1.Name = "TextBox1"
Me.TextBox1.Size = New System.Drawing.Size(188, 20)
Me.TextBox1.TabIndex = 2
Me.TextBox1.Text = "someStringValue "
'
'DateTimePicker1
'
Me.DateTimePicker1.Location = New System.Drawing.Point(12, 92)
Me.DateTimePicker1.Name = "DateTimePicker1"
Me.DateTimePicker1.Size = New System.Drawing.Size(188, 20)
Me.DateTimePicker1.TabIndex = 3
Me.DateTimePicker1.Value = New Date(1987, 5, 22, 0, 0, 0, 0)
'
'txtResult
'
Me.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
						Or System.Windows.Forms.AnchorStyles.Right)
Me.txtResult.Location = New System.Drawing.Point(12, 164)
Me.txtResult.Multiline = True
Me.txtResult.Name = "txtResult"
Me.txtResult.Size = New System.Drawing.Size(292, 68)
Me.txtResult.TabIndex = 4
Me.txtResult.Text = ""
'
'Button1
'
Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
Me.Button1.Location = New System.Drawing.Point(12, 124)
Me.Button1.Name = "Button1"
Me.Button1.Size = New System.Drawing.Size(188, 23)
Me.Button1.TabIndex = 5
Me.Button1.Text = "Generate WHERE clause"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(316, 241)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.txtResult, Me.DateTimePicker1, Me.TextBox1, Me.Label2})
Me.Name = "Form1"
Me.Text = "QuickWhere component demo, Example 1"
Me.ResumeLayout(False)

		End Sub

#End Region

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		With QW
			.Add("SearchField1", "Table1", TypeOfValues.StringType, WhereConditions.BeginsWith, TextBox1)
			.Add("SearchField2", TypeOfValues.DateType, DateTimePicker1)
		End With

	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		txtResult.Text = QW.GetSql
	End Sub

	Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		QW = Nothing
	End Sub

End Class
