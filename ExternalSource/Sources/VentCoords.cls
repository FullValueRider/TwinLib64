VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VentCoords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

    
    Private Enum CoordType
    
        ctX1 = 0
        ctY1
        ctX2
        ctY2
        
    End Enum
    
    Private Type State
        
        X1 As Long
        Y1 As Long
        X2 As Long
        Y2 As Long
        Type As VentLineType
        
    End Type
    
    Private s As State
    
    
    
    Public Function Deb(ByVal ipCoords As String) As VentCoords
        
        With New VentCoords
        
            Set Deb = .ReadyToUseInstance(ipCoords)
            
        End With
        
    End Function
    
    
    Public Function ReadyToUseInstance(ByVal ipCoords As String) As VentCoords
        
        Dim mycoords As Variant
        mycoords = VBA.Split(ipCoords, Char.twComma)
        
        s.X1 = mycoords(ctX1)
        s.Y1 = mycoords(ctY1)
        s.X2 = mycoords(ctX2)
        s.Y2 = mycoords(ctY2)
            
        'Determine any equivalence between X1,X2 and Y1,Y2 so that we
        'can assign a type to the line
        
        Select Case True
            
            Case Is = ((s.X1 = s.X2) And (s.Y1 = s.Y2)):            s.Type = VentLineType.vlPoint
            Case Is = (s.Y1 = s.Y2):                                s.Type = VentLineType.vlHorizontal
            Case Is = (s.X1 = s.X2):                                s.Type = VentLineType.vlVertical
            Case Else:                                              s.Type = VentLineType.vlDiagonal
            
        End Select
        
        Set ReadyToUseInstance = Me
        
    End Function
    
    
    Public Function GetLineType() As VentLineType
        GetLineType = s.Type
    End Function
    
    
    Public Function GetCoords() As Lyst
        
        Dim myVentLine As Lyst
        Set myVentLine = Lyst.Deb
        
        Dim myXCoord As Long
        Dim myYCoord As Long
        Dim myYStep As Long
        myYStep = IIf(s.Y1 <= s.Y2, 1, -1)
        Dim myXStep As Long
        mystep = IIf(s.X1 <= s.X2, 1, -1)
        
        Select Case s.Type
            
            Case vlPoint
            
                myVentLine.Add fmt.Txt("{0},{1}", s.X1, s.Y1)
                
                
            Case vlHorizontal
            
                For myXCoord = s.X1 To s.X2 Step myXStep
                    
                    myVentLine.Add fmt.Txt("{0},{1}", myXCoord, s.Y1)
                    
                Next
            
                
            Case vlVertical
            
                For myYCoord = s.Y1 To s.Y2 Step myYStep
                    
                    myVentLine.Add fmt.Txt("{0},{1}", s.X1, myYCoord)
                    
                Next
            
            
            Case vlDiagonal
            
                myYCoord = s.Y1
                For myXCoord = s.X1 To s.X2 Step myXStep
                    
                    myVentLine.Add fmt.Txt("{0},{1}", myXCoord, myYCoord)
                    myYCoord = myYCoord + myYStep
                    
                Next
            
        End Select
        
        GetCoords = myVentLine
        
    End Function


    
   
