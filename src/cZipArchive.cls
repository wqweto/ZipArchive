VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cZipArchive (c) 2017 by wqweto@gmail.com
'
' A single-class pure VB6 library for zip archives management
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cZipArchive"

#Const ImplLogging = DEBUG_MODE
#Const ImplCompress = ZIP_COMPRESS
#Const ImplDecompress = ZIP_EXTRACT
#Const ImplUseShared = ZIP_USESHARED

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, Cancel As Boolean)
Event Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
#If ImplDecompress Then
    Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
#End If ' ImplDecompress

'=========================================================================
' Thunk data
'=========================================================================

' Auto-generated on 8.4.2017 21:56:36, CodeSize=6240, DataSize=1048, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "i1QkDItEJARTVYtsJBiLWDhWi3QkFAPWi0UAi8pXM/8rzjvWG9L30iPRdieNZCQAD7YOjXYBM8FHi8jB6ASD4Q8zBIuLyMHoBIPhDzMEizv6ct1fXolFAF1bwhAAzMzMVot0JAhXagyLRhz/0Iv4V4k36GkJAACLThxqFP/Rx0AQAAAAAMdADAAAAACJRwiLx19ewhAAzMzMzMzMzMzMzFaLdCQIV4s+/3YIi0ck/9D/dgSLRyT/0ItHJFb/0F9ewhAAzMzMzMzMzMzMzMzMzItEJAxWi3QkDFeLfCQMhcB0DVD/dgT/Nv836BH/////dhiNRgz/dhT/dhBQjUYIUP92BP82V+iFCQAAX17CEAD/dCQE6EcQAADCEADMzMzMVot0JAiLTgyNRgxXiz6FyXQMO04EdAdQVujyEQAAi04QjUYQhcl0DDtOCHQHUFbo3BEAAIN+FACNRhR0B1BW6MwRAACNRgRQVujCEQAAjUYIUFbouBEAAItHJFb/0F9e" & _
    "whAAzMzMzMzMzMzMzMzMzItEJAhTVYtsJAxWV41YDFONeAhX/3AE/zBV6IEJAACLTCQci/CFyXQYhfZ0FIsHhcB0DFH/M1D/dQDoMf7//4vGX15dW8IQAMzMzMzMzMzMuAEAAADCDADMzMzMzMzMzFaLdCQIg8r/i4YEgAQAD7dMxgKNBMZmO8p0CQ+/wWaJFMbrEItABIP4/3QIZomURgiABACLhgSABACLTCQQiUzGBIuGBIAEAGaJVMYCjRROi44EgAQAD7eCCIAEAGaJBM6LhgSABAAPvwzGD7fAZomCCIAEAIP5/3QMZouGBIAEAGaJRM4Ci44EgAQAikQkDIiEDgAABACLhgSABABAJf9/AACJhgSABABewgwAzMzMgewoAwAAU4ucJDQDAABVi6wkNAMAAFYz9leLbQSLvCREAwAAObUMgAYAD47WAAAAi40MgAYAi8ErxgPHg/gED4yYAAAAO/F9Cg+2hC4IgAYA6wiLxivBD7YEGI1WAohE" & _
    "JBSNQv87wX0KD7aELgmABgDrCYvGK8EPtkQYAYhEJBU70X0KD7aELgqABgDrCYvGK8EPtkQYAohEJBaNQgE7wX0KD7aELguABgDrCYvGK8EPtkQYA4hEJBeLRCQUaci9pzUewekQUVBV6IL+//9GO7UMgAYAD4xV////6ymLzju1DIAGAH0fjZUIgAYAjUkAioQNCIAGAI1SAYhC/0E7jQyABgB86im1DIAGAIX/D44mAgAAg/8ED4yVAQAAaQO9pzUeM9LB6BAPv7RFCIAEAIP+/w+EeQEAAIoDiEQkE5CLjQSABACLxivBvwCAAAAl/38AACv4i8/32XkWi4UEgAQAA8El/38AAIqEKAAABADrA4oEGThEJBMPhYcAAACNQQGFwHkXi4UEgAQAQAPBJf9/AACKhCgAAAQA6wSKRBkBOEMBdWCNQQKFwHkZi4UEgAQAg8ACA8El/38AAIqEKAAABADrBIpEGQI4QwJ1N41BA4XAeRmLhQSABACDwAMD" & _
    "wSX/fwAAioQoAAAEAOsEikQZAzhDA3UOiXyUGEI7lCRIAwAAfQ4Pv3T1AIP+/w+FLP///4lUJBSF0g+OkgAAAIu8JEQDAAC+BAAAADv+fmcz/zPAhdJ+WIoMHohMJBOQi1SEGIvOK8p5GYuNBIAEACvKA86B4f9/AACKjCkAAAQA6wOKDBk4TCQTdQWJVLwYR0A7RCQUfMiF/3QUi9dGi7wkRAMAAIlUJBQ793yi6weLvCREAwAAVv90JBz/tCREAwAA6NgOAACF9g+OggAAAOs7i4wkPAMAAIoDPI+LEYtxCA+2yItCMHcJD7ZECDBqCOsND7YECGoJjQRFAQAAAFBWUuh4EgAAvgEAAACLvCREAwAAg/8EfBZpA72nNR7B6BBQD7YDUFXoQ/z//+sVi40MgAYAigOIhCkIgAYA/4UMgAYATkNPhfZ/yYm8JEQDAACF/w+P2v3//19eXVuBxCgDAADCEADMzMzMzMzMzMyB7DwDAABTi5wkTAMAAFWL" & _
    "rCRMAwAAVleLvCRQAwAAM/aLfwQ5twyABgAPjtYAAACLjwyABgCLwSvGA8OD+AQPjJgAAAA78X0KD7aEPgiABgDrCIvGK8EPtgQojVYCiEQkGI1C/zvBfQoPtoQ+CYAGAOsJi8YrwQ+2RCgBiEQkGTvRfQoPtoQ+CoAGAOsJi8YrwQ+2RCgCiEQkGo1CATvBfQoPtoQ+C4AGAOsJi8YrwQ+2RCgDiEQkG4tMJBhpwb2nNR7B6BBQUVfoMvv//0Y7twyABgAPjFX////rKYvOO7cMgAYAfR+NlwiABgCNSQCKhA8IgAYAjVIBiEL/QTuPDIAGAHzqKbcMgAYAM8AzyYlEJCCJTCQkiUQkGIXbD44xAgAAg/sED4zyAQAAaUUAvac1HsdEJBQAAAAAwegQD7+URwiABACD+v8PhM8BAACKRQCLnwSABACIRCQTjUkAi8K+AIAAACvDJf9/AAAr8IvO99l5EY0EGSX/fwAAioQ4AAAEAOsDigQpOEQkE3Ve" & _
    "jUEBhcB5E41DAQPBJf9/AACKhDgAAAQA6wSKRCkBOEUBdTuNQQKFwHkTjUMCA8El/38AAIqEOAAABADrBIpEKQI4RQKLRCQUdRiJdIQsQIlEJBQ7hCRcAwAAfRPrBItEJBQPvxTXg/r/D4Vh////hcAPjhQBAAC+BAAAADm0JFgDAAAPjoMAAACNpCQAAAAAM9LHRCQoAAAAAIXAfm6L/4X2eRaLhwSABAADxiX/fwAAipw4AAAEAOsDihwui0yULIvGK8F5GIvGK8EDhwSABAAl/38AAIqEOAAABADrA4oEKDrYi0QkKHUJiUyELECJRCQoQjtUJBR8poXAdA5GiUQkFDu0JFgDAAB8hItcJCCF235gjUMBO/B+UYuEJFADAACLEItYCItEJBg8jw+2yItCMHcJD7ZECDBqCOsND7YECGoJjQRFAQAAAFBTUugtDwAAD7ZFAItMJCyJTCQkiXQkIIlEJBjpnQAAAItEJCRTUOsli0QkLIlEJCQPtkQk" & _
    "E4l0JCCJRCQY636LTCQki1wkIIXbfjtTUf+0JFgDAADo+woAADPAjXP/iUQkIIX2f1yLnCRYAwAAi0wkJIXbD4/P/f//X15dW4HEPAMAAMIQAIuMJFADAACKRQA8j4sRi3EID7bIi0IwdwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFZS6HcOAAC+AQAAAIucJFgDAACD+wR8GGlFAL2nNR7B6BBQD7ZFAFBX6ED4///rFouPDIAGAIpFAIiEOQiABgD/hwyABgBORUuF9n/GiZwkWAMAAOle////zMzMzMxWV4t8JAxoEIAGAIsHi0Ac/9CL8IX2dQVfXsIEAIl3BI1OAroAgAAAg8j/6wONSQDHQQL/////jUkIZolB+GaJQfaD6gF16Y2+CIAEALkAgAAA86tfiZYEgAQAjUIBiZYMgAYAXsIEAMzMzMzMzMzMVleLfCQMuQ8AAAC+AECRYbgAKJFh86WLTCQQg8BgX17HAQAQkWGJQQTHQQgAMJFh" & _
    "x0EMGDSRYcIIAMzMzMzMzFNVi2wkIFZXi3wkFGoDjUUCUIt3CIsfVlPHBgAAAADHRggAAAAAx0YEAAAAAOg+DQAAg3wkLAD/dCQw/3QkIP90JCBXdAfoxff//+sF6A77//9qB2oAVlPoEw0AAIXtdBiLThCFyXQRuAgAAAArwVBqAFZT6PcMAACLRCQgiw5fiQiLRCQgi04EXl2JCLgBAAAAW8IgAMzMzMzMzMzMzMxRU1VWi3QkFFdoAAEAAIseiVwkHItDHP/Qi2wkIImGpIEAAMeGrIEAAAABAADHhqiBAAAAAAAAhe1/DYO+nAEAAAAPjrYFAACDvpwBAAAYjb6cAQAAfSqLVCQc6wONSQCF7X4ZD7YCTYsPQtPgCYaYAQAAjUEIiQeD+Bh844lUJByLRhiFwHUJx0YYAQAAAOumg/gBdXOLF4P6Aw+MXQUAAIuGmAEAAIPC/dHoi8iJF8HoAomGmAEAAIPhA3Udi8rHRhgKAAAAg+EHK9HT6IkX" & _
    "iYaYAQAA6Vz///+D+QF1GItGBIlGDItGCIlGEMdGGAYAAADpP////4P5Ag+FNv///4lOGOku////g/gCdW6LF4P6Dg+M5QQAAIuOmAEAAIvBg+AfwekFBQEBAADHRiwAAAAAiUYgi8GD4B/B6QVAx0YYAwAAAIlGJIvBg+APwekEg8AEiY6YAQAAiUYojULyiQczwIlGRIlGSIlGTIlGUGaJRlSIRlbpu/7//4P4A3VvOQcPjHUEAACLRiw7Rih9MYM/A3wsi0s0i0YsipaYAQAAgOIHD7YEAYhUMET/RiyLRiyDB/3BrpgBAAADO0YofM+LRiw7RigPhWn+//9qE41GRFBW6NYJAACJRhTHRhgEAAAAx0YsAAAAAOlH/v//g/gED4XiAAAAi0Yki04gA8E5Rix8P1GNRldQVuifCQAA/3YkiUYMi0Ygg8BXA8ZQVuiKCQAAjX4UiUYQV1bozQUAAMcHAAAAAMdGGAYAAADp8v3///92FI2GmAEAAFdQ" & _
    "6CsGAACLyIP5/w+EngMAAIP5/g+EYgMAAIP5EH0Pi0YsiEwwV/9GLOm6/f//dQe4AgAAAOsPM8CD+REPlcCNBIUDAAAAiUYwM8CD+RIPlMCNBMUDAAAAiUY0g/kQdRuLRiyFwH4UD7ZEMFaJRjzHRhgFAAAA6W39//8zwMdGGAUAAACJRjzpXP3//4P4BXVcix+LTjA72Q+MEQMAAIuGmAEAALoBAAAA0+Ir2UqJHyPQ0+gDVjSJhpgBAACF0n4ci0YkA0Ygi04sO8h9D4pGPEqIRDFX/0YshdJ/5ItcJBjHRhgEAAAA6fv8//+D+AYPhZwAAAD/dgyNXgxXjYaYAQAAUOgoBQAAg/j/D4SdAgAAg/j+D4RdAgAAPQABAAB9EFBW6AgEAACLXCQY6bb8//91QIsDx0YYAQAAADtGBHQNU1boaAQAAMcDAAAAAItGEI1+EItcJBg7RggPhIb8//9XVuhIBAAAxwcAAAAA6XT8//+LXCQYPR4BAAAPjWX8"
Private Const STR_THUNK2 As String = _
    "///HRhgHAAAAiUYc6Vb8//+D+Ad1VItGHC0BAQAAjQxAi0MojRyID79LAjkPD4z7AQAAi5aYAQAAuAEAAADT4EgjwgNDBIlGOA+/QwIpB4pLAotcJBjT6omWmAEAAMdGGAgAAADp/fv//4P4CHU6/3YQjYaYAQAAV1DoMQQAAIP4/w+EpgEAAIP4/g+EagEAAIP4Hg+NYQEAAMdGGAkAAACJRhzpvvv//4P4CQ+FjwAAAItGHIsXjQxAi0MsjQSID79IAolEJCCJTCQQO9EPjFoBAACLhpgBAAC7AQAAACtUJBDT44tMJCBLI9gDWQSJF4pJAtPog344AImGmAEAAMdGGAYAAAB0LI2kJAAAAACLhqCBAAD/TjgrwyX/fwAAD7aEMKABAABQVuiBAgAAg344AHXb/044i1wkGOkm+///g/gKdTGLF4P6EA+M3QAAAIuOmAEAAA+3wYlGQI1C8MHpEIkHiY6YAQAAx0YYCwAAAOnw+v//g/gLdT+LB4P4" & _
    "EA+MpwAAAIuOmAEAAIPA8A+30YkHgfL//wAAi0ZAwekQiY6YAQAAO8J1UPfYG8CD4AtAiUYY6az6//+D+AwPhaP6//+DPwh8ZQ+2hpgBAABQVujZAQAAgwf4wa6YAQAACINGQP8PhXz6///HRhgBAAAA6XD6//+LXCQY/7akgQAAi0Mk/9CLRCQkx4akgQAAAAAAAF9exwAAAAAAi0QkIF1bxwAAAAAAM8BZwhQAi0wkJLgBAAAAi5akgQAAX4kRi5aogQAAi0wkJF5dW4kRWcIUAMzMzMzMzMzMzMzMzMyB7CABAABTVou0JCwBAABXaLCBAACLRhz/0IvYjXwkDLgICAgIuSQAAABoIAEAAIkz86u4CQkJCY28JKAAAAC5HAAAAPOrjUQkEMeEJBABAAAHBwcHUFPHhCQcAQAABwcHB8eEJCABAAAHBwcHx4QkJAEAAAcHBwfHhCQoAQAABwcHB8eEJCwBAAAHBwcHx4QkMAEAAAgICAjHhCQ0AQAA" & _
    "CAgICOjXBAAAiUMEjUQkDGogUFPHRCQYBQUFBcdEJBwFBQUFx0QkIAUFBQXHRCQkBQUFBcdEJCgFBQUFx0QkLAUFBQXHRCQwBQUFBcdEJDQFBQUF6IcEAABfiUMIi8Nex0MYAAAAAMdDFAAAAADHQxAAAAAAx0MMAAAAAMeDmAEAAAAAAADHg5wBAAAAAAAAx4OggQAAAAAAAFuBxCABAADCBADMzMzMzMzMzMzMzMxTi1wkDFaLdCQMi4aggQAAiw6InDCgAQAAi4aggQAAQCX/fwAAiYaggQAAi4aogQAAO4asgQAAfB+NBEWAAAAAUP+2pIEAAImGrIEAAItBIP/QiYakgQAAi4aogQAAi46kgQAAiBwB/4aogQAAXlvCCADMzMzMzMzMzMzMzMzMzFOLXCQIVYtsJBCLA4lEJAyF7XUIXYPI/1vCCABXi30Ahf90S1Yz9jk3fCPrA41JAItHBI0E8IPABIM4AHQHUFPou////0Y7N37mi0QkFP93" & _
    "BItAJP/Qi0QkFMdHBAAAAABXi0Ak/9DHRQAAAAAAXl9dM8BbwggAzMzMzMzMzMzMi0QkDFOLXCQMVosIi0AEizNXi3wkEIsXI8qNBMgPtgg7zn8mD7YI0+or8WaDeAL/dSyLQASFwHQaiwiLQAQjyo0EyA+2CDvOftpfXoPI/1vCDABfXrj+////W8IMAIkXiTMPv0ACX15bwgwAzMzMzIPsCItEJAxViyiLQAiJRCQQi0QkGIlsJASFwA+OLQEAAFNWVz0EAQAAfge+AgEAAOsMi/A9AgEAAH4DjXD9K8a7HQAAAIlEJCSDz/+LRSiL6IlEJBSNBDuZK8LR+I0MQDt0jQR9BIvY6+s7dI0IfgSL+Ovhi1wkHI0EQI08hQAAAAC5FwEAAA+3BC8D/YtsJBBmO8EPv8iLRTB/DA+2hEgA/v//agfrBw+2RAGoaghQU1XoKQMAAA+3RwJmhcB0DSt3BJhQVlNV6BMDAACLRSyDzv+LXCQgvx4AAACL6Iv/" & _
    "jQQ3mSvC0fiNDEA7XI0EfQSL+OvrO1yNCH4Ei/Dr4Yt8JByNBECNNIUAAAAAA/WLbCQQagUPvw6LRTAPtgTIUFdV6LkCAAAPt0YCZoXAdA+YUIvDK0YEUFdV6KECAACLRCQkhcAPj9n+//9fXltdg8QIwgwAzMzMzMzMzMzMzMxTVYtsJAxWV2oIi0Uc/9CLXCQoi/iLy7gBAAAA0+C+AQAAAItMJCxI0+aJRCQojQT1AAAAAFCLRRz/0IlHBDPJjUb/g8r/iQeFwHgjjWQkAItHBGaJVMgCi0cExgTIAItHBMdEyAQAAAAAQTsPfuEz7TlsJCB+aYt0JBzrA41JAA+2BC47w35Ri0QkGIsEqIvII0wkKDtMJCR1PovL0/gjBzsHfzSLTwRmiWzBAotPBA+2FC4r0400wQ+2DjvKfQKIFot0JBy6AQAAAA+2DC4ry9PiA8I7B37MRTtsJCB8oDP2OTd8Z4tUJCyLbCQci/+LRwSNDPAPtgE7wn5Lg8j/" & _
    "ZolBAotHBI0M8A+2ASvCg/gHfgW4BwAAAFCNBBOIEVCLy4vG0+ALRCQsUP90JCxV/3QkLP90JCzow/7//4tPBItUJCyJRPEERjs3fqOLx19eXVvCHADMzMzMzMzMzMzMgewEBQAAM8Az0lWLrCQUBQAAiVQkBIlEJAyJRCQQiUQkFIlEJBiJRCQciUQkIIlEJCSJRCQoiUQkLIlEJDCJRCQ0iUQkOIlEJDyJRCRAiUQkRFaLtCQUBQAAhe1+GYv/D7YMMP9EjAw70X0Ci9FAO8V87YlUJAgz0jPJ6wONSQCJVAxQA1QMEIPBBAPSg/k8fO5XM/+F7X5JU42cJJQAAAAPtjQ3xwMAAAAAi1S0VI1CAYlEtFSF9nQYM8CNZCQAi8oDwIPhAdH6C8GD7gF18IkDi7QkHAUAAEeDwwQ7/XzAW4tEJAxfg/gJfAW4CQAAAFBqAGoAVf+0JCQFAACNhCSgAAAAUIuEJCgFAAD/MOiM/f//Xl2BxAQFAADCDADM" & _
    "i0QkDFaLdCQMi04Q0+AJRgyLRCQUA8GJRhCD+Ah8PleLfCQMi0YEO0YIfBSNBEWAAAAAUP82iUYIi0cg/9CJBotWBIsOikYMiAQK/0YEwW4MCINGEPiDfhAIfchfXsIQAIBAwCCgYOAQkFDQMLBw8AiISMgoqGjoGJhY2Di4ePgEhETEJKRk5BSUVNQ0tHT0DIxMzCysbOwcnFzcPLx8/AKCQsIiomLiEpJS0jKycvIKikrKKqpq6hqaWto6unr6BoZGxiamZuYWllbWNrZ29g6OTs4urm7uHp5e3j6+fv4BgUHBIaFh4RGRUdExsXHxCYlJySmpaekZmVnZObl5+QWFRcUlpWXlFZVV1TW1dfUNjU3NLa1t7R2dXd09vX39A4NDwyOjY+MTk1PTM7Nz8wuLS8srq2vrG5tb2zu7e/sHh0fHJ6dn5xeXV9c3t3f3D49Pzy+vb+8fn1/fP79//wEBAAADAAAAAwAAAAIBAAAEAAAABAAAAAMBAAAFAAAA" & _
    "BQAAAAQBAAAGAAAABgAAAAUBAAAHAAAABwAAAAYBAAAIAAAACAAAAAcBAAAJAAAACQAAAAgBAAAKAAAACgAAAAkBAQALAAAADAAAAAoBAQANAAAADgAAAAsBAQAPAAAAEAAAAAwBAQARAAAAEgAAAA0BAgATAAAAFgAAAA4BAgAXAAAAGgAAAA8BAgAbAAAAHgAAABABAgAfAAAAIgAAABEBAwAjAAAAKgAAABIBAwArAAAAMgAAABMBAwAzAAAAOgAAABQBAwA7AAAAQgAAABUBBABDAAAAUgAAABYBBABTAAAAYgAAABcBBABjAAAAcgAAABgBBABzAAAAggAAABkBBQCDAAAAogAAABoBBQCjAAAAwgAAABsBBQDDAAAA4gAAABwBBQDjAAAAAQEAAB0BAAACAQAAAgEAABAREgAIBwkGCgULBAwDDQIOAQ8AAAAAAAEAAAABAAAAAQAAAAIAAAACAAAAAgAAAAMAAAADAAAAAwAAAAQAAAAEAAAA" & _
    "BAABAAUAAAAGAAAABQABAAcAAAAIAAAABgACAAkAAAAMAAAABwACAA0AAAAQAAAACAADABEAAAAYAAAACQADABkAAAAgAAAACgAEACEAAAAwAAAACwAEADEAAABAAAAADAAFAEEAAABgAAAADQAFAGEAAACAAAAADgAGAIEAAADAAAAADwAGAMEAAAAAAQAAEAAHAAEBAACAAQAAEQAHAIEBAAAAAgAAEgAIAAECAAAAAwAAEwAIAAEDAAAABAAAFAAJAAEEAAAABgAAFQAJAAEGAAAACAAAFgAKAAEIAAAADAAAFwAKAAEMAAAAEAAAGAALAAEQAAAAGAAAGQALAAEYAAAAIAAAGgAMAAEgAAAAMAAAGwAMAAEwAAAAQAAAHAANAAFAAAAAYAAAHQANAAFgAAAAgAAAAAAAAGQQtx3IIG47rDDZJpBB3Hb0UWtrWGGyTTxxBVAgg7jtRJMP8Oij1taMs2HLsMJkm9TS04Z44gqgHPK9vQ=="
Private Const STR_THUNK_OFFSETS As String = "96|160|208|272|288|400|0|0|0|0|6496|6864|6240|6844|7224"
Private Const STR_THUNK_BUILDDATE As String = "8.4.2017 21:56:36"
' end of generated code

'=========================================================================
' API
'=========================================================================

'--- for FindFirstFile
Private Const MAX_PATH                      As Long = 260
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for CreateFile
Private Const CREATE_ALWAYS                 As Long = 2
Private Const OPEN_EXISTING                 As Long = &H3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for SetFilePointer
Private Const FILE_BEGIN                    As Long = 0
Private Const FILE_CURRENT                  As Long = 1
Private Const FILE_END                      As Long = 2
'--- for MultiByteToWideChar
Private Const CP_OEMCP                      As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for SetFilePointer
Private Const INVALID_SET_FILE_POINTER      As Long = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
#If ImplCompress Then
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
#End If ' ImplCompress
#If ImplDecompress Then
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
#End If ' ImplDecompress
#If Not ImplUseShared Then
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
#End If ' Not ImplUseShared

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_SIG_LOCAL         As Long = &H4034B50
Private Const LNG_SIG_CENTRAL       As Long = &H2014B50
Private Const LNG_SIG_END           As Long = &H6054B50
Private Const LNG_LEN_LOCAL         As Long = 30
Private Const LNG_LEN_CENTRAL       As Long = 46
Private Const LNG_LEN_END           As Long = 22
Private Const LNG_VERMADE           As Long = 20
Private Const LNG_VEREXT            As Long = 20
Private Const LNG_METHOD_DEFLATE    As Long = 8
Private Const LNG_METHOD_STORE      As Long = 0
Private Const LNG_IOBUF_SIZE        As Long = 65536
Private Const ERR_USER_CANCEL       As String = "User cancelled"
Private Const ERR_INIT_COMPRESSOR   As String = "Cannot init deflate compressor"
Private Const ERR_COMPRESSING       As String = "Error compressing"
Private Const ERR_EMPTY_ARCHIVE     As String = "Empty archive"
Private Const ERR_INVALID_ARCHIVE   As String = "Invalid archive"
Private Const ERR_ARCHIVE_NOT_OPEN  As String = "Archive not opened"
Private Const ERR_INVALID_LOCALHDR  As String = "Invalid local header"
Private Const ERR_UNSUPPORTED_METHOD As String = "Unsupported compression method"
Private Const ERR_INIT_DECOMPRESSOR As String = "Cannot init inflate decompressor"
Private Const ERR_READING_ARCHIVE   As String = "Error reading archive"
Private Const ERR_DECOMPRESSING     As String = "Error decompressing"
Private Const ERR_CRC_CHECK         As String = "CRC check failed"

Private m_uRtbl                 As UcsZlibRelocTableType
Private m_sLastError            As String
Private m_bCancel               As Boolean
Private m_lFileCount            As Long
Private m_uFiles()              As UcsFileInfo
Private m_lCurrentFile          As Long
#If ImplDecompress Then
    Private m_sComment          As String
    Private m_vArchiveFile      As Variant
#End If ' ImplDecompress

Private Type UcsFileInfo
    FileName            As String
    Attributes          As Long
    Crc32               As Long
    Size                As Long
    CompressedSize      As Long
    Comment             As String
    LastModified        As Date
    Method              As Long
    Offset              As Long
    SourceFile          As Variant
End Type

Private Type UcsVfsFileType
    Handle              As Long
    Proxy               As Object
    Data                As WIN32_FIND_DATA
    FileName            As String
End Type

Private Type UcsZlibRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Tab            As Long
End Type

Private Type UcsZlibBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
End Type

Private Enum UcsRelocIndexesEnum
    ucsIdx_CompressInit = 0
    ucsIdx_CompressCleanup
    ucsIdx_CompressBlock
    ucsIdx_DecompressInit
    ucsIdx_DecompressCleanup
    ucsIdx_DecompressBlock
    ucsIdx_CalcCrc32
    ucsIdx_MallocImpl
    ucsIdx_ReallocImpl
    ucsIdx_FreeImpl
    ucsIdx_LenCodes
    ucsIdx_DistCodes
    ucsIdx_MirrorBytes
    ucsIdx_LenLenMap
    ucsIdx_Crc32Tab
End Enum

Private Type UcsLocalHeaderType
    Signature       As Long                     'Signature
    VerExt          As Integer                  'version needed to extract
    Flags           As Integer                  'encrypt and compression flags
    Method          As Integer                  'compression method
    FTime           As Integer                  'time last modifies, dos format
    FDate           As Integer                  'date last modifies, dos format
    '--- padding
    Crc32           As Long                     'CRC32 for uncompressed file
    CSize           As Long                     'compressed size
    USize           As Long                     'uncompressed size
    LenFname        As Integer                  'Length filename
    LenExt          As Integer                  'Length for extra field
End Type

Private Type UcsCentralHeaderType
    Signature       As Long                     'Signature
    VerMade         As Integer                  'version made by
    VerExt          As Integer                  'version needed to extract
    Flags           As Integer                  'encrypt and compression flags
    Method          As Integer                  'compression method
    FTime           As Integer                  'time last modifies, dos format
    FDate           As Integer                  'date last modifies, dos format
    '--- padding
    Crc32           As Long                     'CRC32 for uncompressed file
    CSize           As Long                     'compressed size
    USize           As Long                     'uncompressed size
    LenFname        As Integer                  'Length filename
    LenExt          As Integer                  'Length for extra field
    LenCom          As Integer                  'Length for comment field
    DiskStart       As Integer                  'start disk number
    AttribI         As Integer                  'internal file attributes
    '--- padding
    AttribX         As Long                     'external file attributes
    Offset          As Long                     'relative offset of local header
End Type

Private Type UcsEndHeaderType
    Signature       As Long                     'Signature
    DiskNum         As Integer                  'this disk number
    DiskStart       As Integer                  'start disk number
    Entries         As Integer                  'Entries on this disk
    TotEntr         As Integer                  'Number of total entries
    CenSize         As Long                     'size of entire cetral directory
    CenOff          As Long                     'offset of central on starting disk
    LenCom          As Integer                  'Length of comment field
End Type

'=========================================================================
' Properties
'=========================================================================

Property Get SemVersion() As String
    SemVersion = "0.1.2"
End Property

Property Get ThunkBuildDate() As String
    ThunkBuildDate = STR_THUNK_BUILDDATE
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property

Property Get FileInfo(ByVal FileIdx As Long) As Variant
    If FileIdx >= 0 And FileIdx < m_lFileCount Then
        With m_uFiles(FileIdx)
            FileInfo = Array(.FileName, .Attributes, .Crc32, .Size, .CompressedSize, .Comment, .LastModified, .Method, .Offset)
        End With
    End If
End Property

#If ImplDecompress Then

Public Property Get Comment() As String
    Comment = m_sComment
End Property
    
#End If ' ImplDecompress

'=========================================================================
' Methods
'=========================================================================

#If ImplCompress Then

Public Function AddFile( _
            File As Variant, _
            Optional Name As String, _
            Optional Comment As String) As Boolean
    Const FUNC_NAME     As String = "AddFile"
    Dim uFile           As UcsVfsFileType
    
    On Error GoTo EH
    m_sLastError = vbNullString
    If m_lFileCount = 0 Then
        ReDim m_uFiles(0 To 2) As UcsFileInfo
    ElseIf m_lFileCount > UBound(m_uFiles) Then
        ReDim Preserve m_uFiles(0 To 2 * UBound(m_uFiles)) As UcsFileInfo
    End If
    uFile = pvVfsOpen(File)
    With m_uFiles(m_lFileCount)
        If LenB(Name) <> 0 Then
            .FileName = Name
        Else
            .FileName = Mid$(uFile.FileName, InStrRev(uFile.FileName, "\") + 1)
        End If
        .Size = IIf(Right(.FileName, 1) = "\", 0, uFile.Data.nFileSizeLow)
        .Attributes = IIf(Right(.FileName, 1) = "\", vbDirectory, uFile.Data.dwFileAttributes)
        .Comment = Comment
        .LastModified = pvFromFileTime(uFile.Data.ftLastWriteTime)
        AssignVariant .SourceFile, File
    End With
    m_lFileCount = m_lFileCount + 1
    '--- success
    AddFile = True
QH:
    On Error Resume Next
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME, Err.Description
    Resume QH
End Function

Public Function CompressArchive(ArchiveFile As Variant, Optional Comment As String, Optional ByVal Level As Long = 6) As Boolean
    Const FUNC_NAME     As String = "CompressArchive"
    Dim baBuffer()      As Byte
    Dim baNext()        As Byte
    Dim uArchiveFile    As UcsVfsFileType
    Dim uLocal          As UcsLocalHeaderType
    Dim lIdx            As Long
    Dim uFile           As UcsVfsFileType
    Dim lSize           As Long
    Dim lBeginOffset    As Long
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    Dim uCentral        As UcsCentralHeaderType
    Dim uEndHdr         As UcsEndHeaderType
    Dim lEntries        As Long
    
    On Error GoTo EH
    m_sLastError = vbNullString
    m_lCurrentFile = -1
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsCreate(ArchiveFile)
    uLocal.Signature = LNG_SIG_LOCAL
    uLocal.VerExt = LNG_VEREXT
    Level = LimitLong(Level, 0, 9)
    uBuf.Greedy = -(Level <= 4)
    uBuf.MaxMatch = At(Array(0, 1, 4, 8, 24, 8, 16, 32, 64, 200), Level)
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            .Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            On Error GoTo EH_Continue
            uFile = pvVfsOpen(.SourceFile)
            pvToDosDateTime .LastModified, uLocal.FDate, uLocal.FTime
            uLocal.Method = LNG_METHOD_DEFLATE
            uLocal.Crc32 = -1
            uLocal.USize = 0
            uLocal.CSize = 0
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName
            lBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            If (.Attributes And vbDirectory) = 0 Then
                If Level > 0 Then
                    hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
                    End If
                    lSize = pvVfsRead(uFile, VarPtr(baNext(0)), UBound(baNext) + 1)
                    Do
                        Call CopyMemory(baBuffer(0), baNext(0), lSize)
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lSize
                        lSize = pvVfsRead(uFile, VarPtr(baNext(0)), UBound(baNext) + 1)
                        uBuf.Final = -(lSize = 0)
                        lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(uLocal.Crc32), Level) '--- level ignored
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_COMPRESSING
                        End If
                        pvVfsWrite uArchiveFile, uBuf.OutBlock, uBuf.OutLen
                        Call CoTaskMemFree(uBuf.OutBlock)
                        uBuf.OutBlock = 0
                        uLocal.USize = uLocal.USize + uBuf.InLen
                        uLocal.CSize = uLocal.CSize + uBuf.OutLen
                        .CompressedSize = uLocal.CSize
                        RaiseEvent Progress(lIdx, uLocal.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                    Loop While lSize <> 0
                    If hCtx <> 0 Then
                        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                        hCtx = 0
                    End If
                    If uLocal.CSize = 0 And uLocal.USize = 0 Then
                        uLocal.Method = LNG_METHOD_STORE
                    End If
                End If
                '--- if data is incompressible -> store file
                If Level = 0 Or uLocal.CSize > uLocal.USize Then
                    pvVfsSeek uFile, 0, FILE_BEGIN
                    pvVfsSeek uArchiveFile, lBeginOffset, FILE_BEGIN
                    uLocal.Method = LNG_METHOD_STORE
                    uLocal.Crc32 = -1
                    uLocal.USize = 0
                    uLocal.CSize = 0
                    Do
                        RaiseEvent Progress(lIdx, uLocal.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                        lSize = pvVfsRead(uFile, VarPtr(baBuffer(0)), UBound(baBuffer) + 1)
                        If lSize = 0 Then
                            Exit Do
                        End If
                        Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lSize, VarPtr(uLocal.Crc32))
                        pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), lSize
                        uLocal.USize = uLocal.USize + lSize
                        uLocal.CSize = uLocal.CSize + lSize
                    Loop
                End If
            End If
            pvVfsClose uFile
            uLocal.Crc32 = uLocal.Crc32 Xor -1
            lBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName
            pvVfsSeek uArchiveFile, lBeginOffset, FILE_BEGIN
            .Crc32 = uLocal.Crc32
            Debug.Assert .Size = uLocal.USize
            .Size = uLocal.USize
            .CompressedSize = uLocal.CSize
            .Method = uLocal.Method
            If False Then
SkipFile:
                On Error GoTo EH
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                .Offset = -1
                If hCtx <> 0 Then
                    Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                    hCtx = 0
                End If
            End If
            On Error GoTo EH
        End With
    Next
    m_lCurrentFile = -1
    lBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
    uCentral.Signature = LNG_SIG_CENTRAL
    uCentral.VerMade = LNG_VERMADE
    uCentral.VerExt = LNG_VEREXT
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            If .Offset >= 0 Then
                uCentral.Method = .Method
                pvToDosDateTime .LastModified, uCentral.FDate, uCentral.FTime
                uCentral.Crc32 = .Crc32
                uCentral.CSize = .CompressedSize
                uCentral.USize = .Size
                uCentral.AttribX = .Attributes
                uCentral.Offset = .Offset
                pvOutputCentralHeader uArchiveFile, uCentral, .FileName, .Comment
                lEntries = lEntries + 1
            End If
        End With
    Next
    m_lCurrentFile = -1
    If lEntries = 0 Then
        Err.Raise vbObjectError, , ERR_EMPTY_ARCHIVE
    End If
    uEndHdr.Signature = LNG_SIG_END
    uEndHdr.Entries = lEntries
    uEndHdr.TotEntr = lEntries
    uEndHdr.CenSize = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT) - lBeginOffset
    uEndHdr.CenOff = lBeginOffset
    pvOutputEndHeader uArchiveFile, uEndHdr, Comment
    pvVfsSetEof uArchiveFile, "[output_archive]"
    '--- success
    CompressArchive = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uFile
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME, Err.Description) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Public Function AddFromFolder( _
            sFolderAndMask As String, _
            Optional Recursive As Boolean, _
            Optional TargetFolder As String, _
            Optional IncludeEmptyFolders As Boolean) As Boolean
    Const FUNC_NAME     As String = "AddFromFolder"
    Dim lIdx            As Long
    Dim sFolder         As String
    Dim sMask           As String
    Dim lRootOffset     As Long
    Dim cFolders        As Collection
    Dim vElem           As Variant
    Dim sName           As String
    
    On Error GoTo EH
    m_sLastError = vbNullString
    Set cFolders = New Collection
    lIdx = InStrRev(sFolderAndMask, "\")
    If lIdx > 0 Then
        sFolder = Left(sFolderAndMask, lIdx - 1)
    End If
    sMask = Mid$(sFolderAndMask, lIdx + 1)
    If LenB(sFolder) = 0 Then
        sFolder = CurDir$()
    End If
    lRootOffset = Len(sFolder)
    cFolders.Add sFolder
    Do While cFolders.Count > 0
        sFolder = cFolders.Item(1)
        cFolders.Remove 1
        lIdx = 0
        For Each vElem In pvEnumFiles(sFolder, sMask, -vbDirectory)
            If AddFile(vElem, PathCombine(TargetFolder, Mid$(vElem, lRootOffset + 2))) Then
                lIdx = lIdx + 1
                '--- success (entries added)
                AddFromFolder = True
            End If
        Next
        If lIdx = 0 And IncludeEmptyFolders Then
            sName = PathCombine(TargetFolder, Mid$(sFolder, lRootOffset + 2))
            If LenB(sName) <> 0 Then
                If AddFile(sFolder & "\", sName & "\") Then
                    '--- success (entries added)
                    AddFromFolder = True
                End If
            End If
        End If
        If Recursive Then
            For Each vElem In pvEnumFiles(sFolder, "*.*", vbDirectory)
                cFolders.Add vElem
            Next
        End If
    Loop
QH:
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME, Err.Description
    Resume QH
End Function

Private Function pvEnumFiles(sFolder As String, sMask As String, ByVal eAttrib As VbFileAttribute) As Collection
    Dim sFile           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    
    On Error GoTo EH
    Set pvEnumFiles = New Collection
    sFile = PathCombine(sFolder, sMask)
    hFind = FindFirstFile(sFile, uData)
    If hFind = INVALID_HANDLE_VALUE Then
        If Err.LastDllError <> 2 Then
            On Error GoTo 0
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & sFile & ")"
        End If
    Else
        Do
            If eAttrib > 0 And (uData.dwFileAttributes And eAttrib) <> 0 _
                    Or eAttrib < 0 And (uData.dwFileAttributes And -eAttrib) = 0 Then
                sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                If sFile <> "." And sFile <> ".." Then
                    pvEnumFiles.Add PathCombine(sFolder, sFile)
                End If
            End If
        Loop While FindNextFile(hFind, uData) <> 0
        Call FindClose(hFind)
        hFind = 0
    End If
    Exit Function
EH:
    If hFind <> 0 And hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        hFind = 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub pvOutputLocalHeader(uFile As UcsVfsFileType, uHdr As UcsLocalHeaderType, FileName As String)
    Dim baFName()       As Byte
    
    baFName = pvToOemString(Replace(FileName, "\", "/"))
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = 0
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) + VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = LNG_LEN_LOCAL
    pvVfsWrite uFile, VarPtr(uHdr.Signature), VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature)
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32)
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
End Sub

Private Sub pvOutputCentralHeader(uFile As UcsVfsFileType, uHdr As UcsCentralHeaderType, FileName As String, Comment As String)
    Dim baFName()       As Byte
    Dim baComment()     As Byte
    
    baFName = pvToOemString(Replace(FileName, "\", "/"))
    baComment = pvToOemString(Comment)
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = 0
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) + VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) + VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = LNG_LEN_CENTRAL
    pvVfsWrite uFile, VarPtr(uHdr.Signature), VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature)
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32)
    pvVfsWrite uFile, VarPtr(uHdr.AttribX), VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX)
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvOutputEndHeader(uFile As UcsVfsFileType, uHdr As UcsEndHeaderType, Comment As String)
    Dim baComment()     As Byte
    
    baComment = pvToOemString(Comment)
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.LenCom) + 2 - VarPtr(uHdr.Signature) = LNG_LEN_END
    pvVfsWrite uFile, VarPtr(uHdr.Signature), LNG_LEN_END
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvToDosDateTime(dDate As Date, nDate As Integer, nTime As Integer)
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME
    
    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call FileTimeToDosDateTime(uFileTime, VarPtr(nDate), VarPtr(nTime))
End Sub

Private Function pvToOemString(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    If Len(sText) <> 0 Then
        lSize = WideCharToMultiByte(CP_OEMCP, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
        If lSize > 0 Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(CP_OEMCP, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), UBound(baRetVal) + 1, 0, 0)
            pvToOemString = baRetVal
        Else
            pvToOemString = ApiEmptyByteArray()
        End If
    Else
        pvToOemString = ApiEmptyByteArray()
    End If
End Function

#End If ' ImplCompress

#If ImplDecompress Then

Public Function OpenArchive(ArchiveFile As Variant) As Boolean
    Const FUNC_NAME     As String = "OpenArchive"
    Const MAX_END_SEEK  As Long = 10000
    Dim uArchiveFile    As UcsVfsFileType
    Dim uEndHdr         As UcsEndHeaderType
    Dim lIdx            As Long
    Dim baComment()     As Byte
    Dim sArchiveComment As String
    Dim uFiles()        As UcsFileInfo
    Dim uCentral        As UcsCentralHeaderType
    Dim baFName()       As Byte
    
    On Error GoTo EH
    m_sLastError = vbNullString
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(ArchiveFile)
    For lIdx = 0 To MAX_END_SEEK
        If pvVfsSeek(uArchiveFile, -LNG_LEN_END - lIdx, FILE_END) <= 0 Then
            lIdx = MAX_END_SEEK + 1
            Exit For
        End If
        Debug.Assert VarPtr(uEndHdr.LenCom) + 2 - VarPtr(uEndHdr.Signature) = LNG_LEN_END
        pvVfsRead uArchiveFile, VarPtr(uEndHdr), LNG_LEN_END
        If uEndHdr.Signature = LNG_SIG_END And uEndHdr.LenCom = lIdx Then
            Exit For
        End If
    Next
    If lIdx > MAX_END_SEEK Then
        Err.Raise vbObjectError, , ERR_INVALID_ARCHIVE & " (missing ECDR)"
        GoTo QH
    End If
    If uEndHdr.LenCom > 0 Then
        ReDim baComment(0 To uEndHdr.LenCom - 1) As Byte
        pvVfsRead uArchiveFile, VarPtr(baComment(0)), uEndHdr.LenCom
        sArchiveComment = pvFromOemString(baComment)
    End If
    '--- note: redim one more (last ignored)
    ReDim uFiles(0 To uEndHdr.Entries) As UcsFileInfo
    pvVfsSeek uArchiveFile, uEndHdr.CenOff, FILE_BEGIN
    For lIdx = 0 To uEndHdr.Entries - 1
        m_lCurrentFile = lIdx
        With uFiles(lIdx)
            Debug.Assert VarPtr(uCentral.FDate) + 2 - VarPtr(uCentral.Signature) + VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32) + VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX) = LNG_LEN_CENTRAL
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), VarPtr(uCentral.FDate) + 2 - VarPtr(uCentral.Signature)
            pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32)
            pvVfsRead uArchiveFile, VarPtr(uCentral.AttribX), VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX)
            .Method = uCentral.Method
            .LastModified = pvFromDosDateTime(uCentral.FDate, uCentral.FTime)
            .Crc32 = uCentral.Crc32
            .CompressedSize = uCentral.CSize
            .Size = uCentral.USize
            .Attributes = uCentral.AttribX
            .Offset = uCentral.Offset
            Debug.Assert uCentral.LenFname > 0
            If uCentral.LenFname > 0 Then
                ReDim baFName(0 To uCentral.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uCentral.LenFname
                .FileName = Replace(pvFromOemString(baFName), "/", "\")
            Else
                If pvSetError(MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". Entry " & lIdx + 1 & " has no filename") Then
                    GoTo QH
                End If
                .FileName = "[entry " & lIdx + 1 & "]"
            End If
            If uCentral.LenCom > 0 Then
                ReDim baComment(0 To uCentral.LenCom - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baComment(0)), uCentral.LenCom
                .Comment = pvFromOemString(baComment)
            End If
        End With
    Next
    '--- commit archive info to member vars
    AssignVariant m_vArchiveFile, ArchiveFile
    m_lFileCount = uEndHdr.Entries
    m_sComment = sArchiveComment
    m_uFiles = uFiles
    '--- success
    OpenArchive = True
QH:
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME, Err.Description
    Resume QH
End Function

Public Function Extract( _
            TargetFolder As String, _
            Optional Filter As Variant, _
            Optional OutputFile As Variant) As Boolean
    Const FUNC_NAME     As String = "Extract"
    Dim baBuffer()      As Byte
    Dim uArchiveFile    As UcsVfsFileType
    Dim lIdx            As Long
    Dim uLocal          As UcsLocalHeaderType
    Dim baFName()       As Byte
    Dim baExt()         As Byte
    Dim lRead           As Long
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim vFile           As Variant
    Dim uFile           As UcsVfsFileType
    Dim lCrc32          As Long
    Dim bSkip           As Boolean
    Dim sFilter         As String
    
    On Error GoTo EH
    m_sLastError = vbNullString
    m_lCurrentFile = -1
    If IsEmpty(m_vArchiveFile) Then
        Err.Raise vbObjectError, , ERR_ARCHIVE_NOT_OPEN
    End If
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    If Not IsMissing(Filter) And Not IsArray(Filter) Then
        sFilter = IIf(Filter = "*.*", "*", Replace(Filter, "[", "[[]"))
    End If
    '--- touch to open for read
    pvVfsRead uArchiveFile, 0, 0
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            '--- check if current entry is filtered
            If Not IsMissing(Filter) Then
                If IsArray(Filter) Then
                    If Not At(Filter, lIdx) Then
                        GoTo SkipFile
                    End If
                ElseIf LenB(sFilter) <> 0 Then
                    If Not .FileName Like sFilter Then
                        If Not Mid$(.FileName, InStrRev(.FileName, "\") + 1) Like sFilter Then
                            GoTo SkipFile
                        End If
                    End If
                End If
            End If
            On Error GoTo EH_Continue
            '--- read local header
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) + VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = LNG_LEN_LOCAL
            pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature)
            pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32)
            Debug.Assert uLocal.LenFname > 0
            If uLocal.LenFname > 0 Then
                ReDim baFName(0 To uLocal.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uLocal.LenFname
            End If
            If uLocal.LenExt > 0 Then
                ReDim baExt(0 To uLocal.LenExt - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baExt(0)), uLocal.LenExt
            End If
            '--- sanity checks
            If uLocal.Signature <> LNG_SIG_LOCAL Then
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Signature)"
            End If
            If Replace(pvFromOemString(baFName), "/", "\") <> .FileName Then
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Filename)"
            End If
            If .Method <> LNG_METHOD_DEFLATE And .Method <> LNG_METHOD_STORE Then
                Err.Raise vbObjectError, , ERR_UNSUPPORTED_METHOD & " (" & .Method & ")"
            End If
            '--- allow output to (in-memory) vfs proxy object
            If IsMissing(OutputFile) Then
                vFile = PathCombine(TargetFolder, .FileName)
            Else
                AssignVariant vFile, OutputFile
            End If
            bSkip = False
            RaiseEvent BeforeExtract(lIdx, vFile, bSkip, m_bCancel)
            If m_bCancel Then
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            uFile = pvVfsCreate(vFile)
            If .CompressedSize > 0 Then
                lCrc32 = -1
                If .Method = LNG_METHOD_DEFLATE Then
                    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
                    End If
                    lRead = 0
                    Do ' While lRead < .CompressedSize
                        RaiseEvent Progress(lIdx, lRead, .CompressedSize, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                        If lRead >= .CompressedSize Then
                            Exit Do
                        End If
                        lSize = UBound(baBuffer) + 1
                        If lSize > .CompressedSize - lRead Then
                            lSize = .CompressedSize - lRead
                        End If
                        lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSize)
                        If lResult = 0 Then
                            Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                        End If
                        lRead = lRead + lResult
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lResult
                        lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf), VarPtr(lCrc32))
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_DECOMPRESSING
                        End If
                        pvVfsWrite uFile, uBuf.OutBlock, uBuf.OutLen
                        Call CoTaskMemFree(uBuf.OutBlock)
                        uBuf.OutBlock = 0
                    Loop
                Else ' If .Method = LNG_METHOD_STORE Then
                    lRead = 0
                    Do ' While lRead < .CompressedSize
                        RaiseEvent Progress(lIdx, lRead, .CompressedSize, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                        If lRead >= .CompressedSize Then
                            Exit Do
                        End If
                        lSize = UBound(baBuffer) + 1
                        If lSize > .CompressedSize - lRead Then
                            lSize = .CompressedSize - lRead
                        End If
                        lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSize)
                        If lResult = 0 Then
                            Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                        End If
                        lRead = lRead + lResult
                        Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lSize, VarPtr(lCrc32))
                        pvVfsWrite uFile, VarPtr(baBuffer(0)), lResult
                    Loop
                End If
                pvVfsSetEof uFile, .FileName
                If .Crc32 <> (lCrc32 Xor -1) Then
                    Err.Raise vbObjectError, , ERR_CRC_CHECK
                End If
            Else
                pvVfsSetEof uFile, .FileName
                '--- note: Total=1 to prevent division by zero in client code
                RaiseEvent Progress(lIdx, 1, 1, m_bCancel)
                If m_bCancel Then
                    m_sLastError = ERR_USER_CANCEL
                    GoTo QH
                End If
            End If
            '--- success (entries extracted)
            Extract = True
SkipFile:
            On Error GoTo EH
            pvVfsClose uFile
            If hCtx <> 0 Then
                Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
                hCtx = 0
            End If
        End With
    Next
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uArchiveFile
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME, Err.Description) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Private Function pvFromDosDateTime(ByVal nDate As Integer, ByVal nTime As Integer) As Date
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    If nDate <> 0 Or nTime <> 0 Then
        Call DosDateTimeToFileTime(nDate, nTime, uFileTime)
        Call FileTimeToSystemTime(uFileTime, uSysTime)
        With uSysTime
            pvFromDosDateTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
    End If
End Function

Private Function pvFromOemString(baBuffer() As Byte) As String
    If UBound(baBuffer) >= 0 Then
        pvFromOemString = String(UBound(baBuffer) + 1, 0)
        Call MultiByteToWideChar(CP_OEMCP, 0, baBuffer(0), UBound(baBuffer) + 1, ByVal StrPtr(pvFromOemString), Len(pvFromOemString))
    End If
End Function

#End If ' ImplDecompress

'= VFS ===================================================================

Private Function pvVfsOpen(File As Variant) As UcsVfsFileType
    Dim hFind           As Long
    
    If IsObject(File) Then
        pvVfsOpen.FileName = "[stream]"
        Set pvVfsOpen.Proxy = File
        If Not pvVfsOpen.Proxy Is Nothing Then
            pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.Proxy.VfsSetFilePointer(0, FILE_END)
        End If
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(Now)
    Else
        pvVfsOpen.FileName = File
        If Right$(File, 1) <> "\" Then
            hFind = FindFirstFile(File, pvVfsOpen.Data)
        Else
            hFind = FindFirstFile(Left$(File, Len(File) - 1), pvVfsOpen.Data)
        End If
        If hFind <> INVALID_HANDLE_VALUE Then
            Call FindClose(hFind)
        Else
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & File & ")"
        End If
    End If
End Function

Private Function pvVfsCreate(File As Variant) As UcsVfsFileType
    Dim sError          As String
    
    If IsObject(File) Then
        pvVfsCreate.FileName = "[stream]"
        Set pvVfsCreate.Proxy = File
    Else
        pvVfsCreate.FileName = File
        If InStrRev(File, "\") > 0 Then
            If Not MkPath(Left$(File, InStrRev(File, "\") - 1), sError) Then
                Err.Raise vbObjectError, , sError
            End If
        End If
        If Right$(File, 1) <> "\" Then
            pvVfsCreate.Handle = CreateFile(File, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If pvVfsCreate.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & File & ")"
            End If
        End If
    End If
End Function

Private Function pvVfsClose(uFile As UcsVfsFileType)
    If uFile.Handle <> 0 And uFile.Handle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(uFile.Handle)
        uFile.Handle = 0
    End If
    Set uFile.Proxy = Nothing
End Function

Private Function pvVfsRead(uFile As UcsVfsFileType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not uFile.Proxy Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Proxy.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsRead = uFile.Proxy.VfsReadFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = CreateFile(uFile.FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If ReadFile(uFile.Handle, ByVal lPtr, lSize, pvVfsRead, 0) = 0 Then
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Function pvVfsWrite(uFile As UcsVfsFileType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not uFile.Proxy Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Proxy.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsWrite = uFile.Proxy.VfsWriteFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = CreateFile(uFile.FileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If WriteFile(uFile.Handle, ByVal lPtr, lSize, pvVfsWrite, 0) = 0 Then
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Function pvVfsSeek(uFile As UcsVfsFileType, ByVal lPosition As Long, ByVal lMoveMethod As Long) As Long
    If Not uFile.Proxy Is Nothing Then
        pvVfsSeek = uFile.Proxy.VfsSetFilePointer(lPosition, lMoveMethod)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            '--- touch to open for read
            pvVfsRead uFile, 0, 0
        End If
        pvVfsSeek = SetFilePointer(uFile.Handle, lPosition, 0, lMoveMethod)
        If pvVfsSeek = INVALID_SET_FILE_POINTER Then
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Sub pvVfsSetEof(uFile As UcsVfsFileType, sMetaData As String)
    If Not uFile.Proxy Is Nothing Then
        uFile.Proxy.VfsSetEndOfFile sMetaData
    Else
        If SetEndOfFile(uFile.Handle) = 0 Then
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Sub

'= common ================================================================

Private Function pvFromFileTime(uFileTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME

    Call FileTimeToLocalFileTime(uFileTime, uLocalTime)
    Call FileTimeToSystemTime(uLocalTime, uSysTime)
    With uSysTime
        pvFromFileTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Private Function pvToFileTime(dDate As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uLocalTime      As FILETIME

    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uLocalTime)
    Call LocalFileTimeToFileTime(uLocalTime, pvToFileTime)
End Function

Private Function pvSetError(ByVal Source As String, ByVal Description As String) As Boolean
    #If ImplLogging Then
        If m_lCurrentFile >= 0 Then
            Debug.Print Description, m_uFiles(m_lCurrentFile).FileName, Timer
        Else
            Debug.Print Description, Timer
        End If
    #End If ' ImplLogging
    m_sLastError = Description
    RaiseEvent Error(m_lCurrentFile, Source, Description, m_bCancel)
    pvSetError = m_bCancel
End Function

'= zlib thunks ===========================================================

Private Function pvInitRelocTable(uRtbl As UcsZlibRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(ucsIdx_CompressInit)
        .CompressCleanup = lpThunk + vSplit(ucsIdx_CompressCleanup)
        .CompressBlock = lpThunk + vSplit(ucsIdx_CompressBlock)
        .DecompressInit = lpThunk + vSplit(ucsIdx_DecompressInit)
        .DecompressCleanup = lpThunk + vSplit(ucsIdx_DecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(ucsIdx_DecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(ucsIdx_CalcCrc32)
        .MallocImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(ucsIdx_LenCodes)
        .DistCodes = lpThunk + vSplit(ucsIdx_DistCodes)
        .MirrorBytes = lpThunk + vSplit(ucsIdx_MirrorBytes)
        .LenLenMap = lpThunk + vSplit(ucsIdx_LenLenMap)
        .Crc32Tab = lpThunk + vSplit(ucsIdx_Crc32Tab)
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    
    If lpThunk = 0 Then
        baThunk = FromBase64(STR_THUNK1 & STR_THUNK2)
        lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
    End If
    pvGetThunkAddress = lpThunk
End Function

'= shared ================================================================

#If Not ImplUseShared Then

Private Function MkPath(sPath As String, sError As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(sPath)
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1), sError) Then
                Exit Function
            End If
        End If
        If CreateDirectory(sPath, 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError) & " (" & sPath & ")"
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory) = 0 Then
        sError = "File already exists (" & sPath & ")"
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, lSize)
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Function FromBase64(sValue As String) As Byte()
    With VBA.CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .Text = sValue
        FromBase64 = .NodeTypedValue
    End With
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Private Function LimitLong( _
            ByVal lValue As Long, _
            Optional ByVal Min As Long = -2147483647, _
            Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        LimitLong = Min
    ElseIf lValue > Max Then
        LimitLong = Max
    Else
        LimitLong = lValue
    End If
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
    m_lCurrentFile = -1
End Sub
