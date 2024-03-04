VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CONST_INTEGRATOR

    FORWARD_EULER = 0
    SECOND_ORDER_EULER = 1
    VERLET = 2
    VELOCITY_VERLET = 3
    SECOND_ORDER_RUNGE_KUTTA = 4
    THIRD_ORDER_RUNGE_KUTTA = 5
    FORTH_ORDER_RUNGE_KUTTA = 6
    
End Enum

Private Type POINT2D

    X As Double
    Y As Double

End Type

Private Type VECTOR2D

    X As Double
    Y As Double

End Type

Private Type FORCE2D

    Net As VECTOR2D

End Type


Private Type TORQUE2D

    Net As Double

End Type

Private Type PHYSICS2D
    
    Position As POINT2D
    Velocity As VECTOR2D
    Acceleration As VECTOR2D
    Angle As Double
    Angular_Velocity As Double
    Angular_Acceleration As Double
    Force As FORCE2D
    Torque As TORQUE2D
    Mass As Double
    One_Over_Mass As Double
    Inertia As Double
    One_Over_Inertia As Double
    Elasticity As Double
    Momentum As VECTOR2D
    Impulse As VECTOR2D
    Angular_Momentum As VECTOR2D
    Angular_Impulse As VECTOR2D
    Scalar As Double
    Friction_Coefficient As Double
    Drag_Coefficient As Double
    Area As Double
    
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const EARTH_GRAVITY As Double = 9.80665
Private Const AIR_DENSITY As Double = 1.125

Private Const POUNDS_TO_KG As Double = 0.45359237

Private Const ONE_KGF_CM_TO_NEWTONS_CM As Double = 9.80665
Private Const TWO_KGF_CM_TO_NEWTONS_CM As Double = 9.80665 * 2
Private Const THREE_KGF_CM_TO_NEWTONS_CM As Double = 9.80665 * 3
Private Const FOUR_KGF_CM_TO_NEWTONS_CM As Double = 9.80665 * 4

Private Ticks_Per_Second As Currency
Private Start_Time As Currency
Private Milliseconds As Currency

Private Get_Frames_Per_Second As Long
Private Frame_Count As Long

Private Running As Boolean

Private Time As Double
Private Delta_Time As Currency
Private Accumulator As Double
Private Obj As PHYSICS2D
Private Time_Step As Double

Private Neutral_Flag As Boolean
Private Key_State As Long, Key_Flag As Long

Private Flag As Boolean

Private Old_Position As POINT2D

Private Sub Main()

    With Me

        .ScaleMode = 3
        .AutoRedraw = True
        .BackColor = RGB(0, 0, 0)
        .ForeColor = RGB(0, 255, 0)
        
    End With
    
    With Obj
    
        .Mass = 10 * POUNDS_TO_KG
        .One_Over_Mass = 1 / .Mass
        .Position.X = Me.ScaleWidth / 2
        .Position.Y = 0 'Me.ScaleHeight / 2
        .Velocity.X = 3
        .Velocity.Y = 0
        .Acceleration.X = 0
        .Acceleration.Y = 0
        .Force.Net.X = 0
        .Force.Net.Y = Calculate_Gravitational_Force(.Mass, EARTH_GRAVITY) - Calculate_Air_Resistance(.Drag_Coefficient, AIR_DENSITY, .Area, .Velocity.Y)
        .Scalar = 1000
        .Elasticity = 0.7
        .Friction_Coefficient = 0.7
        .Drag_Coefficient = 0.47 '1000 for water
        .Area = 1
    
    End With

    Time_Step = 1 / 1000

    Running = True

    QueryPerformanceFrequency Ticks_Per_Second
    
    Milliseconds = Time
    
    Neutral_Flag = True
    
    Game_Loop

End Sub

Private Sub Game_Loop()

    Do While Running = True
    
        DoEvents
        
        Lock_Framerate 60
        
        Cls
        
        Old_Position.X = Obj.Position.X
        Old_Position.Y = Obj.Position.Y
         
        Delta_Time = Get_Elapsed_Time_Per_Frame
        
        If Delta_Time > 0.25 Then Delta_Time = 0.25
        
        Accumulator = Accumulator + Delta_Time
        
        While (Accumulator >= Time_Step)
            
            Accumulator = Accumulator - Time_Step
            
            Integrate2D Obj, Time_Step, FORTH_ORDER_RUNGE_KUTTA
            
            Time = Time + Time_Step
            
        Wend
        
        Check_Collision Obj
        
        Render

    Loop

End Sub

Private Function Calculate_Normal_Force(m As Double, g As Double) As Double

   Calculate_Normal_Force = m * g

End Function

Private Function Calculate_Frictional_Force(u As Double, N As Double) As Double

    Calculate_Frictional_Force = u * N

End Function

Private Function Calculate_Gravitational_Force(m As Double, g As Double) As Double

    Calculate_Gravitational_Force = m * g
    
End Function

Private Function Calculate_Air_Resistance(C As Double, p As Double, Area As Double, v As Double)
    
    'F = C*p*A*v²/2

    Calculate_Air_Resistance = (C * p * Area * (v * v)) * 0.5

End Function

Private Sub Lock_Framerate(Target_FPS As Long)

    Static Last_Time As Currency

    Dim Current_Time As Currency
    
    Dim FPS As Double
    
    Do

        QueryPerformanceCounter Current_Time
    
        FPS = Ticks_Per_Second / (Current_Time - Last_Time)
    
    Loop While (FPS > Target_FPS)
    
    QueryPerformanceCounter Last_Time

End Sub

Private Function Get_Elapsed_Time_Per_Frame() As Double

    Static Last_Time As Currency

    Static Current_Time As Currency

    QueryPerformanceCounter Current_Time
    
    Get_Elapsed_Time_Per_Frame = (Current_Time - Last_Time) / Ticks_Per_Second
    
    QueryPerformanceCounter Last_Time

End Function

Private Function Get_Elapsed_Seconds() As Double
    
    Dim Last_Time As Currency
    
    Dim Current_Time As Currency

    QueryPerformanceCounter Current_Time
    
    Get_Elapsed_Seconds = (Current_Time - Last_Time) / Ticks_Per_Second
    
    QueryPerformanceCounter Last_Time
    
End Function

Private Function Get_FPS(Optional ByVal Elapsed_Frames As Long = 1) As Long

    Static Last_Time As Currency

    Dim Current_Time As Currency
    
    QueryPerformanceCounter Current_Time
    
    Get_FPS = Int(Elapsed_Frames * Ticks_Per_Second / (Current_Time - Last_Time))
    
    QueryPerformanceCounter Last_Time
    
End Function

Private Sub Integrate2D(Obj As PHYSICS2D, dt As Double, Integrator As CONST_INTEGRATOR)


    Dim Old_Position As POINT2D
    Dim Old_Velocity As POINT2D
    Dim Old_Acceleration As POINT2D
    Dim Old_Angle As Double
    Dim Old_Angular_Velocity As Double
    Dim Old_Angular_Acceleration As Double

    Dim k1 As POINT2D, k2 As POINT2D, k3 As POINT2D, k4 As POINT2D
    Dim l1 As VECTOR2D, l2 As VECTOR2D, l3 As VECTOR2D, l4 As VECTOR2D
    
    Dim m1 As Double, m2 As Double, m3 As Double, m4 As Double
    Dim n1 As Double, n2 As Double, n3 As Double, n4 As Double
    
    With Obj

        .Acceleration.X = .Force.Net.X * .One_Over_Mass
        .Acceleration.Y = .Force.Net.Y * .One_Over_Mass
        
        .Angular_Acceleration = .Torque.Net * .One_Over_Inertia
        
        Select Case Integrator
        
            Case FORWARD_EULER
                
                .Position.X = .Position.X + .Velocity.X * dt * .Scalar
                .Velocity.X = .Velocity.X + .Acceleration.X * dt
            
                .Position.Y = .Position.Y + .Velocity.Y * dt * .Scalar
                .Velocity.Y = .Velocity.Y + .Acceleration.Y * dt
                
                .Angle = .Angle + .Angular_Velocity * dt
                .Angular_Velocity = .Angular_Velocity + .Angular_Acceleration * dt
                
            Case SECOND_ORDER_EULER
            
                .Position.X = .Position.X + .Velocity.X * dt + 0.5 * .Acceleration.X * dt * dt * .Scalar
                .Velocity.X = .Velocity.X + .Acceleration.X * dt
            
                .Position.Y = .Position.Y + .Velocity.Y * dt + 0.5 * .Acceleration.Y * dt * dt * .Scalar
                .Velocity.Y = .Velocity.Y + .Acceleration.Y * dt
                
                .Angle = .Angle + .Angular_Velocity * dt + 0.5 * .Angular_Acceleration * dt * dt
                .Angular_Velocity = .Angular_Velocity + .Angular_Acceleration * dt
            
            Case VERLET
            
                .Velocity.X = .Position.X - Old_Position.X + .Acceleration.X * dt * dt * .Scalar
                Old_Position.X = .Position.X
                .Position.X = .Position.X + .Velocity.X
                
                .Velocity.Y = .Position.Y - Old_Position.Y + .Acceleration.Y * dt * dt * .Scalar
                Old_Position.Y = .Position.Y
                .Position.Y = .Position.Y + .Velocity.Y
                
                .Angular_Velocity = .Angle - Old_Angle + .Angular_Acceleration * dt * dt
                Old_Angle = .Angle
                .Angle = .Angle + .Angular_Velocity
                
            Case VELOCITY_VERLET
                
                Old_Acceleration.X = .Acceleration.X
                .Position.X = .Position.X + .Velocity.X * dt + 0.5 * Old_Acceleration.X * dt * dt * .Scalar
                .Velocity.X = .Velocity.X + 0.5 * (Old_Acceleration.X + .Acceleration.X) * dt
            
                Old_Acceleration.Y = .Acceleration.Y
                .Position.Y = .Position.Y + .Velocity.Y * dt + 0.5 * Old_Acceleration.Y * dt * dt * .Scalar
                .Velocity.Y = .Velocity.Y + 0.5 * (Old_Acceleration.Y + .Acceleration.Y) * dt
            
                Old_Angular_Acceleration = .Angular_Acceleration
                .Angle = .Angle + .Angular_Velocity * dt + 0.5 + Old_Angular_Acceleration * dt * dt
                .Angular_Velocity = .Angular_Velocity + 0.5 * (Old_Angular_Acceleration + .Angular_Acceleration) * dt
            
            Case SECOND_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity + m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                .Position.X = .Position.X + k2.X * .Scalar
                .Position.Y = .Position.Y + k2.Y * .Scalar
                .Velocity.X = .Velocity.X + l2.X
                .Velocity.Y = .Velocity.Y + l2.Y
                
                
                .Angle = .Angle + m2
                .Angular_Velocity = .Angular_Velocity + n2
            
            Case THIRD_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                k3.X = dt * (.Velocity.X - k1.X + 2 * k2.X)
                k3.Y = dt * (.Velocity.Y - k1.Y + 2 * k2.Y)
                l3.X = dt * .Acceleration.X
                l3.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity * m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                m3 = dt * (.Angular_Velocity - m1 + 2 * m2)
                n3 = dt * .Angular_Acceleration
                
                .Position.X = .Position.X + k1.X * 1 / 6 + k2.X * 2 / 3 + k3.X * 1 / 6 * .Scalar
                .Position.Y = .Position.Y + k1.Y * 1 / 6 + k2.Y * 2 / 3 + k3.Y * 1 / 6 * .Scalar
                .Velocity.X = .Velocity.X + l1.X * 1 / 6 + l2.X * 2 / 3 + l3.X * 1 / 6
                .Velocity.Y = .Velocity.Y + l1.Y * 1 / 6 + l2.Y * 2 / 3 + l3.Y * 1 / 6
                
                .Angle = .Angle + m1 * 1 / 6 + m2 * 2 / 3 + m3 * 1 / 6
                .Angular_Velocity = .Angular_Velocity + n1 * 1 / 6 + n2 * 2 / 3 + n3 * 1 / 6
                
            Case FORTH_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                k3.X = dt * (.Velocity.X + k2.X / 2)
                k3.Y = dt * (.Velocity.Y + k2.Y / 2)
                l3.X = dt * .Acceleration.X
                l3.Y = dt * .Acceleration.Y
                
                k4.X = dt * (.Velocity.X + k3.X)
                k4.Y = dt * (.Velocity.Y + k3.Y)
                l4.X = dt * .Acceleration.X
                l4.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity + m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                m3 = dt * (.Angular_Velocity + m2 / 2)
                n3 = dt * .Angular_Acceleration
                
                m4 = dt * (.Angular_Velocity + m3)
                n4 = dt * .Angular_Acceleration
                 
                .Position.X = .Position.X + k1.X / 6 + k2.X / 3 + k3.X / 3 + k4.X / 6 * .Scalar
                .Position.Y = .Position.Y + k1.Y / 6 + k2.Y / 3 + k3.Y / 3 + k4.Y / 6 * .Scalar
                .Velocity.X = .Velocity.X + l1.X / 6 + l2.X / 3 + l3.X / 3 + l4.X / 6
                .Velocity.Y = .Velocity.Y + l1.Y / 6 + l2.Y / 3 + l3.Y / 3 + l4.Y / 6
                
                .Angle = .Angle + m1 / 6 + m2 / 3 + m3 / 3 + m4 / 6
                .Angular_Velocity = .Angular_Velocity + n1 / 6 + n2 / 3 + n3 / 3 + n4 / 6
                
        End Select
        
    End With

End Sub


Private Sub Clear()

    Me.Cls

End Sub

Private Function Check_Collision(Obj As PHYSICS2D)
    
    Const MIN As Double = 0.4
    
    With Obj
    
        Dim I As Boolean
         
        If .Position.X - Old_Position.X > 0 Then
            
            I = True
            
        ElseIf .Position.X - Old_Position.X < 0 Then
        
            I = False
            
        End If
    
        If .Position.Y >= 175 Then

            .Position.Y = 175
            .Velocity.Y = .Velocity.Y * -.Elasticity
            
            If .Velocity.Y >= 0 Then
            
                If .Velocity.Y <= MIN Then
                
                    .Velocity.Y = 0
                    .Force.Net.Y = 0
                
                End If
                
            Else
            
                If .Velocity.Y >= -MIN Then
                
                    .Velocity.Y = 0
                    .Force.Net.Y = 0
                
                End If
            
            End If

            If I = True Then
            
                .Force.Net.X = -Calculate_Frictional_Force(.Friction_Coefficient, Calculate_Normal_Force(.Mass, EARTH_GRAVITY))
            
                If .Velocity.X <= MIN Then
                
                    .Velocity.X = 0
                    .Force.Net.X = 0
                
                End If
            
            Else
            
                .Force.Net.X = Calculate_Frictional_Force(.Friction_Coefficient, Calculate_Normal_Force(.Mass, EARTH_GRAVITY))
                
                If .Velocity.X >= -MIN Then
                
                    .Velocity.X = 0
                    .Force.Net.X = 0
                   
                End If
                
            End If
                
            
        Else
        
            .Force.Net.X = 0
            
        End If
        
        If .Position.X >= 275 Then

            .Position.X = 275
            .Velocity.X = .Velocity.X * -.Elasticity
            
            .Force.Net.Y = EARTH_GRAVITY + Calculate_Frictional_Force(.Friction_Coefficient, Calculate_Normal_Force(.Mass, EARTH_GRAVITY)) - Calculate_Air_Resistance(.Drag_Coefficient, AIR_DENSITY, .Area, .Velocity.Y)
            
        Else
        
            .Force.Net.Y = Calculate_Gravitational_Force(.Mass, EARTH_GRAVITY) - Calculate_Air_Resistance(.Drag_Coefficient, AIR_DENSITY, .Area, .Velocity.Y)
            
        End If
        
        If .Position.X <= 30 Then
            
            .Position.X = 30
            .Velocity.X = .Velocity.X * -.Elasticity
            
            .Force.Net.Y = EARTH_GRAVITY + Calculate_Frictional_Force(.Friction_Coefficient, Calculate_Normal_Force(.Mass, EARTH_GRAVITY)) - Calculate_Air_Resistance(.Drag_Coefficient, AIR_DENSITY, .Area, .Velocity.Y)
            
        Else
        
            .Force.Net.Y = Calculate_Gravitational_Force(.Mass, EARTH_GRAVITY) - Calculate_Air_Resistance(.Drag_Coefficient, AIR_DENSITY, .Area, .Velocity.Y)
            
        End If
        
    End With

End Function

Private Sub Render()
    
    On Error Resume Next
    
    Dim X As Double, Y As Double
    
    X = Obj.Position.X
    Y = Obj.Position.Y
    
    Circle (X, Y), 15, RGB(0, 255, 0)
    
    Line (0, 190)-(Me.ScaleWidth, 190), RGB(0, 255, 0)

End Sub

Private Sub Close_Program()

    Running = False
    
    Unload Me
    
    End
    
End Sub

Private Sub Form_Activate()

    Main

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Running = False

End Sub
