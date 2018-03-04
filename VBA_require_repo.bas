Attribute VB_Name = "VBA_require_repo"
Option Explicit

Function repoRate(initialPrice As Double, _
                    futurePrice As Double, _
                    days As Integer, _
                    Optional daysInYear As Integer = 365) As Double
    repoRate = (futurePrice / initialPrice - 1) * daysInYear / days
End Function

Function repoPrice(initialPrice As Double, _
                    repoRate As Double, _
                    days As Integer, _
                    Optional daysInYear As Integer = 365) As Double
    repoPrice = initialPrice * (1 + repoRate * days / daysInYear)
End Function
