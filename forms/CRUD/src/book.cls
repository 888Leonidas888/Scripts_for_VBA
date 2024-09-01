VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Type data
    id As Integer
    name As String
    author As String
    isbn As String
    editorial As String
    comment As String
    badge As String
    published As Integer
    price As Double
    created As Date
    updated As Date
End Type

Private self As data

Public Property Get id() As Integer
    id = self.id
End Property

Public Property Let id(ByVal value As Integer)
    self.id = value
End Property

Public Property Get name() As String
    name = self.name
End Property

Public Property Let name(ByVal value As String)
    self.name = UCase(value)
End Property

Public Property Get author() As String
    author = self.author
End Property

Public Property Let author(ByVal value As String)
    self.author = UCase(value)
End Property

Public Property Get isbn() As String
    isbn = self.isbn
End Property

Public Property Let isbn(ByVal value As String)

'    If IsISBN(value) Then
        self.isbn = value
'    Else
'        Err.Raise 10000, Description:="ISBN is not valid"
'    End If
'
End Property

Public Property Get editorial() As String
    editorial = self.editorial
End Property

Public Property Let editorial(ByVal value As String)
    self.editorial = UCase(value)
End Property

Public Property Get comment() As String
    comment = self.comment
End Property

Public Property Let comment(ByVal value As String)
    self.comment = UCase(value)
End Property

Public Property Get badge() As String
    badge = self.badge
End Property

Public Property Let badge(ByVal value As String)
    self.badge = UCase(value)
End Property

Public Property Get published() As Integer
    published = self.published
End Property

Public Property Let published(ByVal value As Integer)
    self.published = value
End Property

Public Property Get price() As Double
    price = self.price
End Property

Public Property Let price(ByVal value As Double)
    self.price = value
End Property

Public Property Get created() As Date
    created = self.created
End Property

Public Property Let created(ByVal value As Date)
    self.created = value
End Property
Public Property Get updated() As Date
    updated = self.updated
End Property

Public Property Let updated(ByVal value As Date)
    self.updated = value
End Property

Private Function IsISBN(ByVal isbn As String) As Boolean
    Dim regex As New RegExp
    Dim pattern As String
    
    pattern = "^(?:\d{1,5}-?\d{1,7}-?\d{1,7}-?[\dXx]|\d{3}-?\d-?\d{1,5}-?\d{1,7}-?\d)$"
    
    regex.pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False

    If regex.test(isbn) Then
        IsISBN = True
    Else
        IsISBN = False
    End If
End Function
