Attribute VB_Name = "Function"
Public Sub Forms(ByVal i As Integer)
 Select Case i
    Case 1
    Load frmMotorInsurance
    frmMotorInsurance.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
     Unload DetailedReportMotor
    Unload DetailedReportNonmotor

    
    
    Case 2
    Load frmEditMotor
    frmEditNonMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor


    
    Case 3
    Load frmSearchMotor
    frmSearchMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor

    
    
    Case 4
    Load frmNonMotor
    frmNonMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor
    
    
    
    Case 5
    Load frmEditNonMotor
    frmEditNonMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor

    
    Case 6
    Load frmSearchNonMotor
    frmSearchNonMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor

    
    Case 7
    Load DetailedReportMotor
    DetailedReportMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportNonmotor
    
    
    Case 8
    Load frmFilterReportMotor
    frmFilterReportMotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportNonmotor
    Unload DetailedReportMotor
    
    
    Case 9
    Load DetailedReportNonmotor
    DetailedReportNonmotor.Show
    Unload frmChangePassword
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    
    Case 10
    Load frmFilterReportNonmotor
    frmFilterReportNonmotor.Show
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportNonmotor
    Unload DetailedReportMotor
 
    Case 11
    Load frmChangePassword
    frmChangePassword.Show
    Unload frmDeleteUser
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor
    
    Case 12
    Load frmDeleteUser
    frmDeleteUser.Show
    Unload frmChangePassword
    Unload frmEditMotor
    Unload frmEditNonMotor
    Unload frmFilterReportMotor
    Unload frmFilterReportNonmotor
    Unload frmMotorInsurance
    Unload frmNonMotor
    Unload frmSearchMotor
    Unload frmSearchNonMotor
    Unload DetailedReportMotor
    Unload DetailedReportNonmotor
    
    
 End Select
End Sub
