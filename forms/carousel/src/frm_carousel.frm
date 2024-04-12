VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_carousel 
   Caption         =   "UserForm1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "frm_carousel.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_carousel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private imgsCollection As New Collection
Private indexImgCollection As Integer
'Carrusel de imagenes, las imagenes son tomadas del directorio ".\img\" dentro del proyecto actual

Private Sub lbl_next_Click()

    Call nextImg

End Sub

Private Sub lbl_prev_Click()

    Call prevImg
    
End Sub

Private Sub UserForm_Initialize()

    indexImgCollection = 1
    
    Call createCollecctionImgs
    Me.img_carousel.Picture = LoadPicture(imgsCollection.Item(indexImgCollection))
    
    Call chargeStyleForm
    
End Sub

Private Sub createCollecctionImgs()
    'Extrea las rutas de las imagenes del directorio ".\img\",
    'luego carga esas rutas en el objeto público imgsCollection, para
    'su posterior llamada.

    Dim fso As New Scripting.FileSystemObject
    Dim f As Folder
    Dim imgPath As String
    
    imgPath = ThisWorkbook.Path + "\img\"
    Set f = fso.GetFolder(imgPath)
    
    For Each File In f.Files
         imgsCollection.Add imgPath + File.Name
    Next File
    
    Set fso = Nothing
    Set f = Nothing
    
End Sub
Private Sub nextImg()
    'Carga la siguiente ruta de la imagen en el control img_carousel,
    'las rutas se encuentran dentro de imgsCollection
    
    indexImgCollection = indexImgCollection + 1
    
    If indexImgCollection > imgsCollection.Count Then indexImgCollection = 1
    Me.img_carousel.Picture = LoadPicture(imgsCollection.Item(indexImgCollection))

End Sub
Private Sub prevImg()
    'Carga la anterior ruta de la imagen en el control img_carousel
    'las rutas se encuentran dentro de imgsCollection
    
    indexImgCollection = indexImgCollection - 1
    
    If indexImgCollection = 0 Then indexImgCollection = imgsCollection.Count
    Me.img_carousel.Picture = LoadPicture(imgsCollection.Item(indexImgCollection))
    
End Sub

Private Sub chargeStyleForm()
    'Carga los estilos para el formulario frm_carousel y sus controles internos.
    
    With Me
        .BackColor = RGB(255, 255, 255)
        .Caption = "Carrusel"
        
         With .frm_containerCarousel
            .Caption = Empty
            .SpecialEffect = fmSpecialEffectFlat
            .BackColor = RGB(255, 255, 255)
         End With
         
         With lbl_next
            .BackColor = RGB(255, 255, 255)
         End With
         
         With lbl_prev
            .BackColor = RGB(255, 255, 255)
         End With
         
        With .img_carousel
            .SpecialEffect = fmSpecialEffectFlat
            .BorderColor = RGB(255, 255, 255)
        End With
    End With
    
End Sub

