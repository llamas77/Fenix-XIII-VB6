Attribute VB_Name = "SurfaceDB"
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    Surface As Texture
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Public Sub Delete()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Public Property Get Surface(ByVal fileIndex As Long) As Texture
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                
                Surface.Width = .SurfaceEntry(i).Surface.Width
                Surface.Height = .SurfaceEntry(i).Surface.Height
                Surface.Ptr = .SurfaceEntry(i).Surface.Ptr
                Exit Property
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As Texture
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler
    
    Dim newSurface As SURFACE_ENTRY_DYN
    Dim Image As CoIoImage
        
    With newSurface
        .fileIndex = fileIndex

        Call CoIoImageLoadFromFile(Image, DirGraficos & fileIndex & ".png")

        '.Surface.Surface = Video.CreateTexture2DFromMemory(Image.X, Image.vY, False, 1, TEXTURE_FORMAT_RGBA8, TEXTURE_FLAG_NONE, BGFX.Copy(Image.vData, Image.vX * Image.vY * Image.vComponent))
        
        Call CreateNormalTexture(.Surface.Ptr, Image)
        
        Call CoIoImageFree(Image)
        
        .Surface.Width = Image.X
        .Surface.Height = Image.Y
        
    End With

    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        LoadSurface.Ptr = newSurface.Surface.Ptr
        LoadSurface.Width = newSurface.Surface.Width
        LoadSurface.Height = newSurface.Surface.Height
        
    End With
    
Exit Function

ErrHandler:

End Function
