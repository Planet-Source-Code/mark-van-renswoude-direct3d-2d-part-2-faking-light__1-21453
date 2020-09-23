Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

'// Main DirectX Declarations
Public mDX As New DirectX7
Public mDDraw As DirectDraw7
Public mD3D As Direct3D7
Public mD3DDevice As Direct3DDevice7

'// Vertex Declarations
Public mtlTile(3) As D3DTLVERTEX

'// Surfaces / Textures Declarations
Public msFront As DirectDrawSurface7
Public msBack As DirectDrawSurface7
Public msTexture(1) As DirectDrawSurface7

'// Screen Declarations
Public SCREEN_WIDTH As Long
Public SCREEN_HEIGHT As Long
Public SCREEN_DEPTH As Long
Public SCREEN_BACKCOLOR As Long
Public TILE_WIDTH As Long
Public TILE_HEIGHT As Long
Public LIGHT_MAXSTEPS As Long
Public LIGHT_BEGIN As Long
Public LIGHT_END As Long

'// Other Declarations
Public mbRunning As Boolean
Public pCursor As POINTAPI
Public bTiles(1 To 20, 1 To 15) As Byte
Public Sub InitDX(Width As Long, Height As Long, Depth As Long, DeviceGUID As String)
    Dim ddsd As DDSURFACEDESC2
    Dim caps As DDSCAPS2
    
    '// Create DirectDraw object
    Set mDDraw = mDX.DirectDrawCreate("")
    
    '// Set Cooperative Level (fullscreen, exclusive access)
    mDDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT
    mDDraw.SetDisplayMode Width, Height, Depth, 0, DDSDM_DEFAULT
    
    '// Create primary surface
    ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_3DDEVICE Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    
    Set msFront = mDDraw.CreateSurface(ddsd)
    
    '// Create the backbuffer (used for 3D drawing)
    caps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_3DDEVICE
    Set msBack = msFront.GetAttachedSurface(caps)
    
    '// Create Direct3D
    Set mD3D = mDDraw.GetDirect3D
    
    '// Set the Device
    Set mD3DDevice = mD3D.CreateDevice(DeviceGUID, msBack)
End Sub

Public Sub Start(DeviceGUID As String)
    Dim lX As Long
    Dim lY As Long
    
    '// Show the Form
    frmMain.Show
    DoEvents
    
    '// Screen Settings
    SCREEN_WIDTH = 640
    SCREEN_HEIGHT = 480
    SCREEN_DEPTH = 16
    SCREEN_BACKCOLOR = RGB2DX(0, 0, 0)
    TILE_WIDTH = 32
    TILE_HEIGHT = 32
    LIGHT_MAXSTEPS = 10
    LIGHT_BEGIN = 255
    LIGHT_END = 0
    
    '// Hide the cursor
    Call ShowCursor(0)
    
    '// Initialize DirectX at 640x480x16
    Call InitDX(SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_DEPTH, DeviceGUID)
    
    '// Load Textures and Surfaces
    Set msTexture(0) = CreateTexture(App.Path & "\Texture1.bmp", 32, 32, None)
    Set msTexture(1) = CreateTexture(App.Path & "\Texture2.bmp", 32, 32, None)
    
    '// Set Tiles
    For lY = 1 To 15
        For lX = 1 To 20
            bTiles(lX, lY) = CByte(Rnd(1))
        Next lX
    Next lY
    
    '// Start Running
    Call MainLoop
    
    '// Show the cursor
    Call ShowCursor(1)
    
    '// End it all
    Call Terminate
End Sub


Public Sub ClearDevice()
    '// Clear the device for drawing operations
    Dim rClear(0) As D3DRECT
    
    rClear(0).X2 = SCREEN_WIDTH
    rClear(0).Y2 = SCREEN_HEIGHT
    mD3DDevice.Clear 1, rClear, D3DCLEAR_TARGET, SCREEN_BACKCOLOR, 0, 0
End Sub

Public Sub MainLoop()
    Dim lVertexBrightness As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim lColor As Long
    Dim lX As Long
    Dim lY As Long
    
    mbRunning = True
    
    '// Draw until the program should stop running
    Do While mbRunning
        '// Get position of cursor
        Call GetCursorPos(pCursor)
        
        '// Clear the device
        Call ClearDevice
        
        '// Start the scene
        mD3DDevice.BeginScene
            
            '// Loop trough all tiles
            For lY = 1 To 15
                For lX = 1 To 20
                    '// Create the four vertices which make the tile
                    lVertexBrightness = CalcVertexBrightness(pCursor.x, pCursor.y, (lX - 1) * TILE_WIDTH, (lY - 1) * TILE_HEIGHT)
                    lColor = RGB2DX(lVertexBrightness, lVertexBrightness, lVertexBrightness)
                    Call mDX.CreateD3DTLVertex((lX - 1) * TILE_WIDTH, (lY - 1) * TILE_HEIGHT, 0, 1, lColor, 0, 0, 0, mtlTile(0))
                    
                    lVertexBrightness = CalcVertexBrightness(pCursor.x, pCursor.y, ((lX - 1) * TILE_WIDTH) + TILE_WIDTH, (lY - 1) * TILE_HEIGHT)
                    lColor = RGB2DX(lVertexBrightness, lVertexBrightness, lVertexBrightness)
                    Call mDX.CreateD3DTLVertex(((lX - 1) * TILE_WIDTH) + TILE_WIDTH, (lY - 1) * TILE_HEIGHT, 0, 1, lColor, 0, 1, 0, mtlTile(1))
                    
                    lVertexBrightness = CalcVertexBrightness(pCursor.x, pCursor.y, (lX - 1) * TILE_WIDTH, ((lY - 1) * TILE_HEIGHT) + TILE_HEIGHT)
                    lColor = RGB2DX(lVertexBrightness, lVertexBrightness, lVertexBrightness)
                    Call mDX.CreateD3DTLVertex((lX - 1) * TILE_WIDTH, ((lY - 1) * TILE_HEIGHT) + TILE_HEIGHT, 0, 1, lColor, 0, 0, 1, mtlTile(2))
                    
                    lVertexBrightness = CalcVertexBrightness(pCursor.x, pCursor.y, ((lX - 1) * TILE_WIDTH) + TILE_WIDTH, ((lY - 1) * TILE_HEIGHT) + TILE_HEIGHT)
                    lColor = RGB2DX(lVertexBrightness, lVertexBrightness, lVertexBrightness)
                    Call mDX.CreateD3DTLVertex(((lX - 1) * TILE_WIDTH) + TILE_WIDTH, ((lY - 1) * TILE_HEIGHT) + TILE_HEIGHT, 0, 1, lColor, 0, 1, 1, mtlTile(3))
                    
                    '// Set texture
                    mD3DDevice.SetTexture 0, msTexture(bTiles(lX, lY))
                    
                    '// Draw the tile
                    Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT)
                Next lX
            Next lY
            
        '// End the scene
        mD3DDevice.EndScene
        
        '// Draw Text
        msBack.SetForeColor RGB(255, 255, 255)
        msBack.DrawText 5, 5, "Press '+' and '-' to change the light's size...   (Current: " & CStr(LIGHT_MAXSTEPS) & ")", False
        msBack.DrawText 5, 20, "Press 'A' and 'Z' to change the light's brightness...   (Current: " & CStr(LIGHT_BEGIN) & ")", False
        msBack.DrawText 5, 35, "Press 'S' and 'X' to change the ambient brightness...   (Current: " & CStr(LIGHT_END) & ")", False
        msBack.DrawText 5, 50, "Press 'Esc' to exit...", False
        
        '// Flip
        msFront.Flip Nothing, DDFLIP_WAIT
        DoEvents
    Loop
End Sub

Public Sub Terminate()
    '// Clean up DirectX
    Call mDDraw.RestoreDisplayMode
    Call mDDraw.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    
    Set mD3DDevice = Nothing
    Set mD3D = Nothing
    Set msBack = Nothing
    Set msFront = Nothing
    Set mDDraw = Nothing
    Set mDX = Nothing
    
    Unload frmMain
End Sub


