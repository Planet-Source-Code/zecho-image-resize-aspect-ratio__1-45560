<div align="center">

## Image Resize / Aspect Ratio


</div>

### Description

If you want to load an image and resize it maintaining the original aspect ratio, this is simple, short and works. Coded in VB6 but is most likely compatible with previous versions.

Start with an Image control INSIDE a PictureBox with the default names (Image1, Picture1).

There are two versions of the sub one with and one without comments.

Call this sub when you load a picture and from the Form_Resize() if your Picture1 resizes with the form.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Zecho](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zecho.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zecho-image-resize-aspect-ratio__1-45560/archive/master.zip)





### Source Code

```
<pre>
'~ Commented version -- Uncommented version is below
'~ Simply to show how short and easy this really is
Private Sub Reset_Image()
Image1.Visible = False
If Image1.Picture Then
'~ this is used in case the image changes
'~ if it's not used, the image control is
'~ still the same size as the previous pic
  Image1.Height = Image1.Picture.Height
  Image1.Width = Image1.Picture.Width
  If Image1.Picture.Height > Image1.Picture.Width Then
'~ the Pic is taller than wide
    Image1.Height = Picture1.Height
    Image1.Width = Image1.Width / (Image1.Picture.Height / Image1.Height)
    If Image1.Width > Picture1.Width Then
'~ If the PictureBox isn't square, the pic still may be larger than it
      Image1.Width = Picture1.Width
      Image1.Height = Image1.Picture.Height / (Image1.Picture.Width / Image1.Width)
    End If
  End If
  If Image1.Picture.Width > Image1.Picture.Height Then
'~ Image is wider than tall
    Image1.Width = Picture1.Width
    Image1.Height = Image1.Height / (Image1.Picture.Width / Image1.Width)
    If Image1.Height > Picture1.Height Then
      Image1.Height = Picture1.Height
      Image1.Width = Image1.Picture.Width / (Image1.Picture.Height / Image1.Height)
    End If
  End If
'~ Center Image1 within Picture1
  Image1.Left = (Picture1.Width / 2) - (Image1.Width / 2)
  Image1.Top = (Picture1.Height / 2) - (Image1.Height / 2)
  Image1.Visible = True
End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~ UN-Commented version of the same sub
'~ Don't use both at the same time
Private Sub Reset_Image()
Image1.Visible = False
If Image1.Picture Then
  Image1.Height = Image1.Picture.Height
  Image1.Width = Image1.Picture.Width
  If Image1.Picture.Height > Image1.Picture.Width Then
    Image1.Height = Picture1.Height
    Image1.Width = Image1.Width / (Image1.Picture.Height / Image1.Height)
    If Image1.Width > Picture1.Width Then
      Image1.Width = Picture1.Width
      Image1.Height = Image1.Picture.Height / (Image1.Picture.Width / Image1.Width)
    End If
  End If
  If Image1.Picture.Width > Image1.Picture.Height Then
    Image1.Width = Picture1.Width
    Image1.Height = Image1.Height / (Image1.Picture.Width / Image1.Width)
    If Image1.Height > Picture1.Height Then
      Image1.Height = Picture1.Height
      Image1.Width = Image1.Picture.Width / (Image1.Picture.Height / Image1.Height)
    End If
  End If
  Image1.Left = (Picture1.Width / 2) - (Image1.Width / 2)
  Image1.Top = (Picture1.Height / 2) - (Image1.Height / 2)
  Image1.Visible = True
End If
End Sub
</pre>
```

