VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerWin64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LongInteger) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LongInteger) As Long
    
Private Type LongInteger
    First32Bits As Long
    Second32Bits As Long
End Type

Private Type TimerAttributes
    CounterInitial As Double
    CounterNow As Double
    PerformanceFrequency As Double
End Type

Private Const MaxValue_32Bits = 4294967296#
Private this As TimerAttributes
    
    Private Sub Class_Initialize()
        PerformanceFrequencyLet
    End Sub
    
        Private Sub PerformanceFrequencyLet()
            Dim TempFrequency As LongInteger
            QueryPerformanceFrequency TempFrequency
            this.PerformanceFrequency = ParseLongInteger(TempFrequency)
        End Sub
    
    Public Sub TimerSet()
        Dim TempCounterIntital As LongInteger
        QueryPerformanceCounter TempCounterIntital
        this.CounterInitial = ParseLongInteger(TempCounterIntital)
    End Sub
    
    Public Function CheckQuarterSecondPassed() As Boolean
        CounterNowLet
        If ((this.CounterNow - this.CounterInitial) / this.PerformanceFrequency) >= 0.25 Then
            CheckQuarterSecondPassed = True
        Else
            CheckQuarterSecondPassed = False
        End If
    End Function

    Public Function CheckFiveSecondsPassed() As Boolean
        CounterNowLet
        If ((this.CounterNow - this.CounterInitial) / this.PerformanceFrequency) >= 10 Then
            CheckFiveSecondsPassed = True
        Else
            CheckFiveSecondsPassed = False
        End If
    End Function
    
    Public Sub PrintTimeElapsed()
        CounterNowLet
        If CounterInitalIsSet = True Then
            Dim TimeElapsed As Double
            TimeElapsed = (this.CounterNow - this.CounterInitial) / this.PerformanceFrequency
            Debug.Print Format(TimeElapsed, "0.000000"); " seconds elapsed "
                    
            Dim TicksElapsed As Double
            TicksElapsed = (this.CounterNow - this.CounterInitial)
            Debug.Print Format(TicksElapsed, "#,##0"); " ticks"
        End If
    End Sub

        Private Function CounterNowLet()
            Dim TempTimeNow As LongInteger
            QueryPerformanceCounter TempTimeNow
            this.CounterNow = ParseLongInteger(TempTimeNow)
        End Function

        Private Function CounterInitalIsSet() As Boolean
            If this.CounterInitial = 0 Then
                MsgBox "Counter Initial Not Set"
                CounterInitalIsSet = False
            Else
                CounterInitalIsSet = True
            End If
        End Function
    
        Private Function ParseLongInteger(ByRef LongInteger As LongInteger) As Double
            Dim First32Bits As Double
            First32Bits = LongInteger.First32Bits
                        
            Dim Second32Bits As Double
            Second32Bits = LongInteger.Second32Bits
                        
            If First32Bits < 0 Then First32Bits = First32Bits + MaxValue_32Bits
            If Second32Bits < 0 Then Second32Bits = First32Bits + MaxValue_32Bits
                        
            ParseLongInteger = First32Bits + (MaxValue_32Bits * Second32Bits)
        End Function

