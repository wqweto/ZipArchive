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
#Const ImplUseApiForBase64 = False

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, Cancel As Boolean)
Event Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
#If ImplDecompress Then
    Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    Event ExtractComplete(ByVal FileIdx As Long, File As Variant)
#End If ' ImplDecompress

'=========================================================================
' Thunk data
'=========================================================================

' Auto-generated on 9.4.2017 19:12:30, CodeSize=6336, DataSize=984, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "UYtEJAhTi1wkEFWLbCQYVleLeDgD64tEJCSJbCQQiwD2wwN0HDvddBQPthNDD7bIM9HB6AgzBJf2wwN16IlcJByL1cdEJBgAAAAAK9OD4vyNDBqDwgPB6gI7yxvJ99EjyolMJCB0eYvpjaQkAAAAAIsbi8vB6QgPtvGLyMHpCA+2yTPxi8vB6RAPttGLyMHpEIu0twAIAAAPtskz0YvIwekYM7SXAAQAAIvTweoYM9EPtsgPtsOLXCQcM8iDwwQzNJeJXCQcM7SPAAwAAItMJBiLxkGJTCQYO811lItsJBCLzTP2K8s76xvt99Uj6XQWD7YTjVsBD7bIRjPRwegIMwSXO/V16otMJCRfXl2JAVtZwhAAzMzMzMzMzMzMzMzMVot0JAhXagyLRhz/0Iv4V4k36AkJAACLThxqFP/Rx0AQAAAAAMdADAAAAACJRwiLx19ewhAAzMzMzMzMzMzMzFaLdCQIV4s+/3YIi0ck/9D/dgSLRyT/0ItHJFb/0F9e" & _
    "whAAzMzMzMzMzMzMzMzMzItEJAxWi3QkDFeLfCQMhcB0DVD/dgT/Nv836FH+////dhyNRgz/dhj/dhT/dhBQjUYIUP92BP82V+giCQAAX17CEADMzMzMzMzMzMzMzMzM/3QkBOjnDwAAwhAAzMzMzFaLdCQIi04MjUYMV4s+hcl0DDtOBHQHUFboghEAAItOEI1GEIXJdAw7Tgh0B1BW6GwRAACDfhQAjUYUdAdQVuhcEQAAjUYEUFboUhEAAI1GCFBW6EgRAACLRyRW/9BfXsIQAMzMzMzMzMzMzMzMzMyLRCQIU1WLbCQMVleNWAxTjXgIV/9wBP8wVeghCQAAi0wkHIvwhcl0GIX2dBSLB4XAdAxR/zNQ/3UA6GH9//+Lxl9eXVvCEADMzMzMzMzMzFaLdCQIg8r/i4YEgAQAD7dMxgKNBMZmO8p0CQ+/wWaJFMbrEItABIP4/3QIZomURgiABACLhgSABACLTCQQiUzGBIuGBIAEAGaJVMYCjRRO" & _
    "i44EgAQAD7eCCIAEAGaJBM6LhgSABAAPvwzGD7fAZomCCIAEAIP5/3QMZouGBIAEAGaJRM4Ci44EgAQAikQkDIiEDgAABACLhgSABABAJf9/AACJhgSABABewgwAzMzMg+wMU1WLbCQgVleLfCQgM/aLfwSJfCQYObcMgAYAD47iAAAAi48MgAYAi8ErxgPFg/gED4ycAAAAi1QkJDvxfQoPtoQ+CIAGAOsIi8YrwQ+2BBCNXgKIRCQUjUP/O8F9Cg+2hD4JgAYA6wmLxivBD7ZEEAGIRCQVO9l9Cg+2hD4KgAYA6wmLxivBD7ZEEAKIRCQWjUMBO8F9Cg+2hD4LgAYA6wmLxivBD7ZEEAOIRCQXi0QkFGnIvac1HsHpEFFQV+iK/v//Rju3DIAGAA+MUf///+sxi847twyABgB9J42XCIAGAOsJjaQkAAAAAIv/ioQPCIAGAI1SAYhC/0E7jwyABgB86im3DIAGAIXtD471AQAA6wqNpCQAAAAAjUkA" & _
    "i1wkJIP9BA+MWQEAAIuHAAAEADPtiYcAgAQAaQO9pzUewegQD7+URwiABACD+v8PhDEBAACLjwSABACLwivBvgCAAAAl/38AACvwK86B4f9/AACLhDkAAAQAOwN1Dom0rxCABgBFO2wkLH0JD78U14P6/3XAhe0PjukAAACLRCQoOUQkMH4GiUQkMOsEi0QkML4EAAAAO8Z+cY2bAAAAAIoEHjPJM9uIRCQThe1+WJCLlI8QgAYAi8YrwnkYi4cEgAQAK8IDxiX/fwAAioQ4AAAEAOsLi3wkJIoEOIt8JBg4RCQTdQiJlJ8QgAYAQ0E7zXy9g/sBfg+LRCQwRovri1wkJDvwfJmLXCQki2wkKDv1fT2LlxCABgCLzivKjSwahcl5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrA4oEGTgEKXUIRkE7dCQofNSLbCQoVv+3EIAGAP90JCjooA4AAIX2fn3rOItMJCCKAzyPixGLcQgPtsiLQjB3CQ+2RAgw" & _
    "agjrDQ+2BAhqCY0ERQEAAABQVlLoRxIAAL4BAAAAi2wkKIP9BHwWaQO9pzUewegQUA+2A1BX6HX8///rFYuPDIAGAIoDiIQPCIAGAP+HDIAGAE5DTYX2f8mJbCQoiVwkJIXtD48X/v//X15dW4PEDMIUAMzMzMzMzMzMzMzMzMyD7BRTVYtsJChWV4t8JCgz9ot/BIl8JCA5twyABgAPjuIAAACLjwyABgCLwSvGA8WD+AQPjJwAAACLVCQsO/F9Cg+2hD4IgAYA6wiLxivBD7YEEI1eAohEJBiNQ/87wX0KD7aEPgmABgDrCYvGK8EPtkQQAYhEJBk72X0KD7aEPgqABgDrCYvGK8EPtkQQAohEJBqNQwE7wX0KD7aEPguABgDrCYvGK8EPtkQQA4hEJBuLRCQYaci9pzUewekQUVBX6Gr7//9GO7cMgAYAD4xR////6zGLzju3DIAGAH0njZcIgAYA6wmNpCQAAAAAi/+KhA8IgAYAjVIBiEL/QTuP" & _
    "DIAGAHzqKbcMgAYAM8AzyYlEJBSJTCQYiUQkHIXtD44JAgAA6wqNpCQAAAAAjUkAi1wkLIP9BA+MwAEAAIuHAAAEADPtiYcAgAQAaQO9pzUewegQD7+URwiABACD+v8PhJgBAACLjwSABACLwivBvgCAAAAl/38AACvwK86B4f9/AACLhDkAAAQAOwN1Dom0rxCABgBFO2wkNH0JD78U14P6/3XAhe0PjkwBAACLVCQwi0QkODvCfgaLwolEJDi+BAAAADvGfneKBB4zyTPbiEQkE4XtfmDrB42kJAAAAACLlI8QgAYAi8YrwnkYi4cEgAQAK8IDxiX/fwAAioQ4AAAEAOsLi3wkLIoEOIt8JCA4RCQTdQiJlJ8QgAYAQ0E7zXy9g/sBfg+LRCQ4Rovri1wkLDvwfJGLXCQsi1QkMDvyfTmLlxCABgCLzivKjSwahcl5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrA4oEGTgEKXUIRkE7dCQwfNSLbCQU" & _
    "he1+Xo1FATvwfk+LRCQoixCLaAiLRCQcPI8PtsiLQjB3CQ+2RAgwagjrDQ+2BAhqCY0ERQEAAABQVVLoIA8AAA+2A4uPEIAGAIlMJBiJdCQUiUQkHOmQAAAAi0QkGFVQ6yWLhxCABgCJRCQYD7YDiXQkFIlEJBzrcYtMJBiLbCQUhe1+MlVR/3QkMOjwCgAAM8CNdf+JRCQUhfZ/UotsJDCLTCQYhe0PjwP+//9fXl1bg8QUwhQAi0wkKIoDPI+LEYtxCA+2yItCMHcJD7ZECDBqCOsND7YECGoJjQRFAQAAAFBWUuh2DgAAvgEAAACLbCQwg/0EfBZpA72nNR7B6BBQD7YDUFfopPj//+sVi48MgAYAigOIhA8IgAYA/4cMgAYATkNNhfZ/yYlsJDCJXCQs6Wr////MzMzMzMzMzMxWV4t8JAxosI8GAIsHi0Ac/9CL8IX2dQVfXsIEAIl3BI1OAroAgAAAg8j/6wONSQDHQQL/////jUkIZolB+GaJ" & _
    "QfaD6gF16Y2+CIAEALkAgAAA86tfiZYEgAQAjUIBiZYMgAYAXsIEAMzMzMzMzMzMVleLfCQMuQ8AAAC+AFDYZ7hgKNhn86WLTCQQg8BgX17HAQAQ2GeJQQTHQQigQNhnx0EMeETYZ8IIAMzMzMzMzFOLXCQIVldoAIAAAIs7i3MIi0cc/9CJBjPAOUQkJGoDD5XAx0YIAIAAAIPAAsdGBAAAAABQVlfoNA0AAIN8JCgA/3QkMP90JDD/dCQg/3QkIFN0B+gX+P//6wXoMPv//2oHagBWV+gFDQAAg3wkJAB0GItOEIXJdBG4CAAAACvBUGoAVlfo5gwAAItEJByLDl+JCItEJByLTgReW4kIuAEAAADCJADMzMzMzMzMzMzMUVNVVot0JBRXaACAAACLHolcJByLQxz/0ItsJCCJhqSBAADHhqyBAAAAgAAAx4aogQAAAAAAAIXtfw2DvpwBAAAAD462BQAAg76cAQAAGI2+nAEAAH0qi1QkHOsDjUkA" & _
    "he1+GQ+2Ak2LD0LT4AmGmAEAAI1BCIkHg/gYfOOJVCQci0YYhcB1CcdGGAEAAADrpoP4AXVzixeD+gMPjF0FAACLhpgBAACDwv3R6IvIiRfB6AKJhpgBAACD4QN1HYvKx0YYCgAAAIPhByvR0+iJF4mGmAEAAOlc////g/kBdRiLRgSJRgyLRgiJRhDHRhgGAAAA6T////+D+QIPhTb///+JThjpLv///4P4AnVuixeD+g4PjOUEAACLjpgBAACLwYPgH8HpBQUBAQAAx0YsAAAAAIlGIIvBg+AfwekFQMdGGAMAAACJRiSLwYPgD8HpBIPABImOmAEAAIlGKI1C8okHM8CJRkSJRkiJRkyJRlBmiUZUiEZW6bv+//+D+AN1bzkHD4x1BAAAi0YsO0YofTGDPwN8LItLNItGLIqWmAEAAIDiBw+2BAGIVDBE/0Ysi0Ysgwf9wa6YAQAAAztGKHzPi0YsO0YoD4Vp/v//ahONRkRQVujGCQAAiUYUx0YY" & _
    "BAAAAMdGLAAAAADpR/7//4P4BA+F4gAAAItGJItOIAPBOUYsfD9RjUZXUFbojwkAAP92JIlGDItGIIPAVwPGUFboegkAAI1+FIlGEFdW6L0FAADHBwAAAADHRhgGAAAA6fL9////dhSNhpgBAABXUOgbBgAAi8iD+f8PhJ4DAACD+f4PhGIDAACD+RB9D4tGLIhMMFf/Rizpuv3//3UHuAIAAADrDzPAg/kRD5XAjQSFAwAAAIlGMDPAg/kSD5TAjQTFAwAAAIlGNIP5EHUbi0YshcB+FA+2RDBWiUY8x0YYBQAAAOlt/f//M8DHRhgFAAAAiUY86Vz9//+D+AV1XIsfi04wO9kPjBEDAACLhpgBAAC6AQAAANPiK9lKiR8j0NPoA1Y0iYaYAQAAhdJ+HItGJANGIItOLDvIfQ+KRjxKiEQxV/9GLIXSf+SLXCQYx0YYBAAAAOn7/P//g/gGD4WcAAAA/3YMjV4MV42GmAEAAFDoGAUAAIP4/w+EnQIA"
Private Const STR_THUNK2 As String = _
    "AIP4/g+EXQIAAD0AAQAAfRBQVugIBAAAi1wkGOm2/P//dUCLA8dGGAEAAAA7RgR0DVNW6FgEAADHAwAAAACLRhCNfhCLXCQYO0YID4SG/P//V1boOAQAAMcHAAAAAOl0/P//i1wkGD0eAQAAD41l/P//x0YYBwAAAIlGHOlW/P//g/gHdVSLRhwtAQEAAI0MQItDKI0ciA+/SwI5Dw+M+wEAAIuWmAEAALgBAAAA0+BII8IDQwSJRjgPv0MCKQeKSwKLXCQY0+qJlpgBAADHRhgIAAAA6f37//+D+Ah1Ov92EI2GmAEAAFdQ6CEEAACD+P8PhKYBAACD+P4PhGoBAACD+B4PjWEBAADHRhgJAAAAiUYc6b77//+D+AkPhY8AAACLRhyLF40MQItDLI0EiA+/SAKJRCQgiUwkEDvRD4xaAQAAi4aYAQAAuwEAAAArVCQQ0+OLTCQgSyPYA1kEiReKSQLT6IN+OACJhpgBAADHRhgGAAAAdCyNpCQAAAAA" & _
    "i4aggQAA/044K8Ml/38AAA+2hDCgAQAAUFbogQIAAIN+OAB12/9OOItcJBjpJvv//4P4CnUxixeD+hAPjN0AAACLjpgBAAAPt8GJRkCNQvDB6RCJB4mOmAEAAMdGGAsAAADp8Pr//4P4C3U/iweD+BAPjKcAAACLjpgBAACDwPAPt9GJB4Hy//8AAItGQMHpEImOmAEAADvCdVD32BvAg+ALQIlGGOms+v//g/gMD4Wj+v//gz8IfGUPtoaYAQAAUFbo2QEAAIMH+MGumAEAAAiDRkD/D4V8+v//x0YYAQAAAOlw+v//i1wkGP+2pIEAAItDJP/Qi0QkJMeGpIEAAAAAAABfXscAAAAAAItEJCBdW8cAAAAAADPAWcIUAItMJCS4AQAAAIuWpIEAAF+JEYuWqIEAAItMJCReXVuJEVnCFADMzMzMzMzMzMzMzMzMgewgAQAAU1aLtCQsAQAAV2iwgQAAi0Yc/9CL2I18JAy4CAgICLkkAAAAaCABAACJ" & _
    "M/OruAkJCQmNvCSgAAAAuRwAAADzq41EJBDHhCQQAQAABwcHB1BTx4QkHAEAAAcHBwfHhCQgAQAABwcHB8eEJCQBAAAHBwcHx4QkKAEAAAcHBwfHhCQsAQAABwcHB8eEJDABAAAICAgIx4QkNAEAAAgICAjoxwQAAIlDBI1EJAxqIFBTx0QkGAUFBQXHRCQcBQUFBcdEJCAFBQUFx0QkJAUFBQXHRCQoBQUFBcdEJCwFBQUFx0QkMAUFBQXHRCQ0BQUFBeh3BAAAX4lDCIvDXsdDGAAAAADHQxQAAAAAx0MQAAAAAMdDDAAAAADHg5gBAAAAAAAAx4OcAQAAAAAAAMeDoIEAAAAAAABbgcQgAQAAwgQAzMzMzMzMzMzMzMzMU4tcJAxWi3QkDIuGoIEAAIsOiJwwoAEAAIuGoIEAAEAl/38AAImGoIEAAIuGqIEAADuGrIEAAHwaA8BQ/7akgQAAiYasgQAAi0Eg/9CJhqSBAACLhqiBAACLjqSBAACI" & _
    "HAH/hqiBAABeW8IIAMzMzFOLXCQIVYtsJBCLA4lEJAyF7XUIXYPI/1vCCABXi30Ahf90S1Yz9jk3fCPrA41JAItHBI0E8IPABIM4AHQHUFPou////0Y7N37mi0QkFP93BItAJP/Qi0QkFMdHBAAAAABXi0Ak/9DHRQAAAAAAXl9dM8BbwggAzMzMzMzMzMzMi0QkDFOLXCQMVosIi0AEizNXi3wkEIsXI8qNBMgPtgg7zn8mD7YI0+or8WaDeAL/dSyLQASFwHQaiwiLQAQjyo0EyA+2CDvOftpfXoPI/1vCDABfXrj+////W8IMAIkXiTMPv0ACX15bwgwAzMzMzIPsCItEJAxViyiLQAiJRCQQi0QkGIlsJASFwA+OLQEAAFNWVz0EAQAAfge+AgEAAOsMi/A9AgEAAH4DjXD9K8a7HQAAAIlEJCSDz/+LRSiL6IlEJBSNBDuZK8LR+I0MQDt0jQR9BIvY6+s7dI0IfgSL+Ovhi1wkHI0EQI08hQAA" & _
    "AAC5FwEAAA+3BC8D/YtsJBBmO8EPv8iLRTB/DA+2hEgA/v//agfrBw+2RAGoaghQU1XoKQMAAA+3RwJmhcB0DSt3BJhQVlNV6BMDAACLRSyDzv+LXCQgvx4AAACL6Iv/jQQ3mSvC0fiNDEA7XI0EfQSL+OvrO1yNCH4Ei/Dr4Yt8JByNBECNNIUAAAAAA/WLbCQQagUPvw6LRTAPtgTIUFdV6LkCAAAPt0YCZoXAdA+YUIvDK0YEUFdV6KECAACLRCQkhcAPj9n+//9fXltdg8QIwgwAzMzMzMzMzMzMzMxTVYtsJAxWV2oIi0Uc/9CLXCQoi/iLy7gBAAAA0+C+AQAAAItMJCxI0+aJRCQojQT1AAAAAFCLRRz/0IlHBDPJjUb/g8r/iQeFwHgjjWQkAItHBGaJVMgCi0cExgTIAItHBMdEyAQAAAAAQTsPfuEz7TlsJCB+aYt0JBzrA41JAA+2BC47w35Ri0QkGIsEqIvII0wkKDtMJCR1PovL0/gj" & _
    "BzsHfzSLTwRmiWzBAotPBA+2FC4r0400wQ+2DjvKfQKIFot0JBy6AQAAAA+2DC4ry9PiA8I7B37MRTtsJCB8oDP2OTd8Z4tUJCyLbCQci/+LRwSNDPAPtgE7wn5Lg8j/ZolBAotHBI0M8A+2ASvCg/gHfgW4BwAAAFCNBBOIEVCLy4vG0+ALRCQsUP90JCxV/3QkLP90JCzow/7//4tPBItUJCyJRPEERjs3fqOLx19eXVvCHADMzMzMzMzMzMzMgewEBQAAM8Az0lWLrCQUBQAAiVQkBIlEJAyJRCQQiUQkFIlEJBiJRCQciUQkIIlEJCSJRCQoiUQkLIlEJDCJRCQ0iUQkOIlEJDyJRCRAiUQkRFaLtCQUBQAAhe1+GYv/D7YMMP9EjAw70X0Ci9FAO8V87YlUJAgz0jPJ6wONSQCJVAxQA1QMEIPBBAPSg/k8fO5XM/+F7X5JU42cJJQAAAAPtjQ3xwMAAAAAi1S0VI1CAYlEtFSF9nQYM8CNZCQA" & _
    "i8oDwIPhAdH6C8GD7gF18IkDi7QkHAUAAEeDwwQ7/XzAW4tEJAxfg/gJfAW4CQAAAFBqAGoAVf+0JCQFAACNhCSgAAAAUIuEJCgFAAD/MOiM/f//Xl2BxAQFAADCDADMi0QkDFaLdCQMi04Q0+AJRgyLRCQUA8GJRhCD+Ah8OVeLfCQMi0YIOUYEfA8DwFD/NolGCItHIP/QiQaLVgSLDopGDIgECv9GBMFuDAiDRhD4g34QCH3NX17CEABVi+yLAIBAwCCgYOAQkFDQMLBw8AiISMgoqGjoGJhY2Di4ePgEhETEJKRk5BSUVNQ0tHT0DIxMzCysbOwcnFzcPLx8/AKCQsIiomLiEpJS0jKycvIKikrKKqpq6hqaWto6unr6BoZGxiamZuYWllbWNrZ29g6OTs4urm7uHp5e3j6+fv4BgUHBIaFh4RGRUdExsXHxCYlJySmpaekZmVnZObl5+QWFRcUlpWXlFZVV1TW1dfUNjU3NLa1t7R2dXd09vX39" & _
    "A4NDwyOjY+MTk1PTM7Nz8wuLS8srq2vrG5tb2zu7e/sHh0fHJ6dn5xeXV9c3t3f3D49Pzy+vb+8fn1/fP79//wEBAAADAAAAAwAAAAIBAAAEAAAABAAAAAMBAAAFAAAABQAAAAQBAAAGAAAABgAAAAUBAAAHAAAABwAAAAYBAAAIAAAACAAAAAcBAAAJAAAACQAAAAgBAAAKAAAACgAAAAkBAQALAAAADAAAAAoBAQANAAAADgAAAAsBAQAPAAAAEAAAAAwBAQARAAAAEgAAAA0BAgATAAAAFgAAAA4BAgAXAAAAGgAAAA8BAgAbAAAAHgAAABABAgAfAAAAIgAAABEBAwAjAAAAKgAAABIBAwArAAAAMgAAABMBAwAzAAAAOgAAABQBAwA7AAAAQgAAABUBBABDAAAAUgAAABYBBABTAAAAYgAAABcBBABjAAAAcgAAABgBBABzAAAAggAAABkBBQCDAAAAogAAABoBBQCjAAAAwgAAABsBBQDDAAAA" & _
    "4gAAABwBBQDjAAAAAQEAAB0BAAACAQAAAgEAABAREgAIBwkGCgULBAwDDQIOAQ8AAAAAAAEAAAABAAAAAQAAAAIAAAACAAAAAgAAAAMAAAADAAAAAwAAAAQAAAAEAAAABAABAAUAAAAGAAAABQABAAcAAAAIAAAABgACAAkAAAAMAAAABwACAA0AAAAQAAAACAADABEAAAAYAAAACQADABkAAAAgAAAACgAEACEAAAAwAAAACwAEADEAAABAAAAADAAFAEEAAABgAAAADQAFAGEAAACAAAAADgAGAIEAAADAAAAADwAGAMEAAAAAAQAAEAAHAAEBAACAAQAAEQAHAIEBAAAAAgAAEgAIAAECAAAAAwAAEwAIAAEDAAAABAAAFAAJAAEEAAAABgAAFQAJAAEGAAAACAAAFgAKAAEIAAAADAAAFwAKAAEMAAAAEAAAGAALAAEQAAAAGAAAGQALAAEYAAAAIAAAGgAMAAEgAAAAMAAAGwAMAAEwAAAAQAAA" & _
    "HAANAAFAAAAAYAAAHQANAAFgAAAAgAAA"
Private Const STR_THUNK_OFFSETS As String = "288|352|400|480|496|608|0|0|0|0|6592|6960|6336|6940"
Private Const STR_THUNK_BUILDDATE As String = "9.4.2017 19:12:30"
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
Private Const CP_UTF8                       As Long = 65001
Private Const CP_OEMCP                      As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for SetFilePointer
Private Const INVALID_SET_FILE_POINTER      As Long = -1
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1

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
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
#If ImplCompress Then
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
#End If ' ImplCompress
#If ImplDecompress Then
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
#End If ' ImplDecompress
#If ImplUseShared = 0 Then
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
#End If ' Not ImplUseShared
#If ImplUseApiForBase64 Then
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
#End If ' ImplUseApiForBase64

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
Private Const ERR_ENTRY_INVALID_SIG As String = "Entry %1 has invalid signature"
Private Const ERR_ENTRY_NO_FILENAME As String = "Entry %1 has no filename"
Private Const ERR_MISSING_ECD_RECORD As String = "Missing end-of-central-directory record"
Private Const STR_BUFFER            As String = "[buffer]"
Private Const STR_STREAM            As String = "[stream]"
Private Const STR_ENTRY             As String = "[entry %1]"

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
    Flags               As Long
    Extra()             As Byte
    SourceFile          As Variant
End Type

Private Type UcsVfsFileType
    Handle              As Long
    Stream              As Object
    BufferArray         As Variant
    BufferBase          As Long
    BufferSize          As Long
    BufferPtr           As Long
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
    Crc32Table          As Long
End Type

Private Type UcsZlibBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
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
End Enum

Private Enum UcsHeaderFlagsEnum
    ucsZcfEncrypted = 2 ^ 0                     ' If set, indicates that the file is encrypted
    ucsZcfCompressOptionMask = 2 ^ 1 Or 2 ^ 2   ' 0 - normal, 1 - maximum, 2 - fast, 3 - super fast
    ucsZcfHasDataDescriptor = 2 ^ 3             ' Values are put in the data descriptor immediately following the compressed data
    ucsZcfUseUtf8 = 2 ^ 11                      ' Language encoding flag (EFS)
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
    SemVersion = "0.1.6"
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
    pvSetError
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
        .Size = IIf(Right$(.FileName, 1) = "\", 0, uFile.Data.nFileSizeLow)
        .Attributes = IIf(Right$(.FileName, 1) = "\", vbDirectory, uFile.Data.dwFileAttributes)
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
    pvSetError
    m_lCurrentFile = -1
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsCreate(ArchiveFile)
    uLocal.Signature = LNG_SIG_LOCAL
    uLocal.VerExt = LNG_VEREXT
    Level = LimitLong(Level, 0, 9)
    uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
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
            If (.Attributes And vbDirectory + vbVolume) = 0 Then
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
                        uBuf.Final = (lSize = 0)
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
                uCentral.Flags = IIf(pvFromOemString(pvToOemString(.FileName)) <> .FileName, ucsZcfUseUtf8, 0)
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
    If pvSetError(MODULE_NAME & "." & FUNC_NAME, Err.Description, CanContinue:=True) Then
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
    pvSetError
    Set cFolders = New Collection
    lIdx = InStrRev(sFolderAndMask, "\")
    If lIdx > 0 Then
        sFolder = Left$(sFolderAndMask, lIdx - 1)
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
    
    baFName = pvToOemString(Replace(FileName, "\", "/"), IIf((uHdr.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, CP_OEMCP))
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = 0
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) + VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = LNG_LEN_LOCAL
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) = 14
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 14
    Debug.Assert VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 16
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
End Sub

Private Sub pvOutputCentralHeader(uFile As UcsVfsFileType, uHdr As UcsCentralHeaderType, FileName As String, Comment As String)
    Dim lCodePage       As Long
    Dim baFName()       As Byte
    Dim baComment()     As Byte
    
    lCodePage = IIf((uHdr.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, CP_OEMCP)
    baFName = pvToOemString(Replace(FileName, "\", "/"), lCodePage)
    baComment = pvToOemString(Comment, lCodePage)
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = 0
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) + VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) + VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = LNG_LEN_CENTRAL
    Debug.Assert VarPtr(uHdr.FDate) + 2 - VarPtr(uHdr.Signature) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 16
    Debug.Assert VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) = 22
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 22
    Debug.Assert VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = 8
    pvVfsWrite uFile, VarPtr(uHdr.AttribX), 8
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

Private Function pvToOemString(sText As String, Optional ByVal lCodePage As Long = CP_OEMCP) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    If Len(sText) <> 0 Then
        lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
        If lSize > 0 Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), UBound(baRetVal) + 1, 0, 0)
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
    Dim sArchiveComment As String
    Dim uFiles()        As UcsFileInfo
    Dim uCentral        As UcsCentralHeaderType
    Dim baFName()       As Byte
    Dim baExtra()       As Byte
    Dim baComment()     As Byte
    Dim lCodePage       As Long
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(ArchiveFile)
    lIdx = pvVfsSeek(uArchiveFile, 0, FILE_END)
    If lIdx >= LNG_LEN_END Then
        For lIdx = 0 To LimitLong(lIdx - LNG_LEN_END, , MAX_END_SEEK)
            pvVfsSeek uArchiveFile, -LNG_LEN_END - lIdx, FILE_END
            Debug.Assert VarPtr(uEndHdr.LenCom) + 2 - VarPtr(uEndHdr.Signature) = LNG_LEN_END
            pvVfsRead uArchiveFile, VarPtr(uEndHdr), LNG_LEN_END
            If uEndHdr.Signature = LNG_SIG_END And uEndHdr.LenCom = lIdx Then
                lIdx = -1
                Exit For
            End If
        Next
    End If
    If lIdx >= 0 Then
        Err.Raise vbObjectError, , ERR_INVALID_ARCHIVE & ". " & ERR_MISSING_ECD_RECORD
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
            Debug.Assert VarPtr(uCentral.FDate) + 2 - VarPtr(uCentral.Signature) = 16
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 16
            Debug.Assert VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32) = 22
            pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 22
            Debug.Assert VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX) = 8
            pvVfsRead uArchiveFile, VarPtr(uCentral.AttribX), 8
            If uCentral.Signature <> LNG_SIG_CENTRAL Then
                pvSetError MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
                GoTo QH
            End If
            .Method = uCentral.Method
            .LastModified = pvFromDosDateTime(uCentral.FDate, uCentral.FTime)
            .Crc32 = uCentral.Crc32
            .CompressedSize = uCentral.CSize
            .Size = uCentral.USize
            .Attributes = uCentral.AttribX
            .Offset = uCentral.Offset
            .Flags = uCentral.Flags
            lCodePage = IIf((.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, CP_OEMCP)
            Debug.Assert uCentral.LenFname > 0
            If uCentral.LenFname > 0 Then
                ReDim baFName(0 To uCentral.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uCentral.LenFname
                .FileName = pvToWinFileName(baFName, lCodePage)
            Else
                If pvSetError(MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_NO_FILENAME, "%1", lIdx + 1), CanContinue:=True) Then
                    GoTo QH
                End If
                .FileName = Replace(STR_ENTRY, "%1", lIdx + 1)
            End If
            If uCentral.LenExt > 0 Then
                ReDim baExtra(0 To uCentral.LenExt - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baExtra(0)), uCentral.LenExt
                .Extra = baExtra
            End If
            If uCentral.LenCom > 0 Then
                ReDim baComment(0 To uCentral.LenCom - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baComment(0)), uCentral.LenCom
                .Comment = pvFromOemString(baComment, lCodePage)
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
            OutputTarget As Variant, _
            Optional Filter As Variant) As Boolean
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
    Dim vFileName       As Variant
    Dim uFile           As UcsVfsFileType
    Dim lCrc32          As Long
    Dim bSkip           As Boolean
    Dim sFilter         As String
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    If IsEmpty(m_vArchiveFile) Then
        Err.Raise vbObjectError, , ERR_ARCHIVE_NOT_OPEN
    End If
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    If VarType(Filter) = vbString Then
        sFilter = IIf(Filter = "*.*", "*", Replace(Replace(Filter, "/", "\"), "[", "[[]"))
    End If
    '--- touch to open for read
    pvVfsRead uArchiveFile, 0, 0
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            '--- check if current entry is filtered
            If IsMissing(Filter) Then
                '--- do nothing
            ElseIf IsArray(Filter) Then
                If Not At(Filter, lIdx) Then
                    GoTo SkipFile
                End If
            ElseIf VarType(Filter) = vbString Then
                If Not .FileName Like sFilter Then
                    If Not Mid$(.FileName, InStrRev(.FileName, "\") + 1) Like sFilter Then
                        GoTo SkipFile
                    End If
                End If
            ElseIf Not lIdx = Filter Then
                GoTo SkipFile
            End If
            On Error GoTo EH_Continue
            '--- read local header
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) + VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = LNG_LEN_LOCAL
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) = 14
            pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), 14
            Debug.Assert VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = 16
            pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), 16
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
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
            End If
            If pvToWinFileName(baFName, .Flags) <> .FileName Then
'                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Filename)"
            End If
            If .Method <> LNG_METHOD_DEFLATE And .Method <> LNG_METHOD_STORE Then
                Err.Raise vbObjectError, , ERR_UNSUPPORTED_METHOD & " (" & .Method & ")"
            End If
            bSkip = False
            If VarType(OutputTarget) = vbString Then
                vFileName = PathCombine(CStr(OutputTarget), .FileName)
                RaiseEvent BeforeExtract(lIdx, vFileName, bSkip, m_bCancel)
            Else
                RaiseEvent BeforeExtract(lIdx, OutputTarget, bSkip, m_bCancel)
            End If
            If m_bCancel Then
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            '--- allow output to buffer, stream object or (customized) filename
            If VarType(OutputTarget) = vbString Then
                uFile = pvVfsCreate(vFileName)
            Else
                If IsArray(OutputTarget) And .Size > 0 Then
                    ReDim OutputTarget(0 To .Size - 1) As Byte
                End If
                uFile = pvVfsCreate(OutputTarget)
            End If
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
                If (.Attributes And vbDirectory + vbVolume) = 0 Then
                    pvVfsSetEof uFile, .FileName
                End If
                '--- note: Total=1 to prevent division by zero in client code
                RaiseEvent Progress(lIdx, 1, 1, m_bCancel)
                If m_bCancel Then
                    m_sLastError = ERR_USER_CANCEL
                    GoTo QH
                End If
            End If
            If VarType(OutputTarget) = vbString Then
                RaiseEvent ExtractComplete(lIdx, vFileName)
            Else
                RaiseEvent ExtractComplete(lIdx, OutputTarget)
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
    If pvSetError(MODULE_NAME & "." & FUNC_NAME, Err.Description, CanContinue:=True) Then
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

Private Function pvToWinFileName(baBuffer() As Byte, ByVal lCodePage As Long) As String
    '--- replaces non-ANSI symbols with underscore too
    pvToWinFileName = Replace(Replace(StrConv(StrConv(pvFromOemString(baBuffer, lCodePage), _
        vbFromUnicode), vbUnicode), "/", "\"), "?", "_")
End Function

#End If ' ImplDecompress

Private Function pvArrPtr(vArray As Variant) As Long
    Const VT_BYREF      As Long = &H4000
    Dim lPtr            As Long
    
    '--- note: unsigned add for LARGEADDRESSAWARE processes
    lPtr = Peek(((VarPtr(vArray) Xor &H80000000) + 8) Xor &H80000000)
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        pvArrPtr = Peek(((lPtr Xor &H80000000) + 12) Xor &H80000000)
    End If
End Function

'= VFS ===================================================================

Private Function pvVfsOpen(File As Variant) As UcsVfsFileType
    Dim hFind           As Long
    
    If IsArray(File) Then
        pvVfsOpen.FileName = STR_BUFFER
        Call CopyMemory(pvVfsOpen.BufferArray, File, 16)
        pvVfsOpen.BufferBase = pvArrPtr(File)
        pvVfsOpen.BufferPtr = pvVfsOpen.BufferBase
        If pvVfsOpen.BufferBase <> 0 Then
            pvVfsOpen.BufferSize = UBound(File) + 1
        End If
        pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.BufferSize
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(Now)
    ElseIf IsObject(File) Then
        pvVfsOpen.FileName = STR_STREAM
        Set pvVfsOpen.Stream = File
        If Not pvVfsOpen.Stream Is Nothing Then
            pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.Stream.VfsSetFilePointer(0, FILE_END)
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
    
    If IsArray(File) Then
        pvVfsCreate.FileName = STR_BUFFER
        Call CopyMemory(pvVfsCreate.BufferArray, File, 16)
        pvVfsCreate.BufferBase = pvArrPtr(pvVfsCreate.BufferArray)
        pvVfsCreate.BufferPtr = pvVfsCreate.BufferBase
        If pvVfsCreate.BufferBase <> 0 Then
            pvVfsCreate.BufferSize = UBound(pvVfsCreate.BufferArray) + 1
        End If
    ElseIf IsObject(File) Then
        pvVfsCreate.FileName = STR_STREAM
        Set pvVfsCreate.Stream = File
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
    Dim vEmpty          As Variant
    
    If Not IsEmpty(uFile.BufferArray) Then
        Call CopyMemory(uFile.BufferArray, vEmpty, 16)
    End If
    If uFile.Handle <> 0 And uFile.Handle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(uFile.Handle)
        uFile.Handle = 0
    End If
    Set uFile.Stream = Nothing
End Function

Private Function pvVfsRead(uFile As UcsVfsFileType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not IsEmpty(uFile.BufferArray) Then
        pvVfsRead = LimitLong(lSize, 0, uFile.BufferSize - (uFile.BufferPtr - uFile.BufferBase))
        Call CopyMemory(ByVal lPtr, ByVal uFile.BufferPtr, pvVfsRead)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsRead
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsRead = uFile.Stream.VfsReadFile(lPtr, lSize)
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
    Const MAX_STEP      As Long = 2 ^ 24 '--- 16MB
    Dim lOffset         As Long
    
    If Not IsEmpty(uFile.BufferArray) Then
        lOffset = uFile.BufferPtr - uFile.BufferBase
        If lOffset + lSize > uFile.BufferSize Then
            If lOffset + lSize <= MAX_STEP Then
                uFile.BufferSize = LimitLong(2 ^ Int(Log(lOffset + lSize) / Log(2) + 1), 8192)
            Else
                uFile.BufferSize = (lOffset + lSize + MAX_STEP - 1) And -MAX_STEP
            End If
            ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + lOffset
        End If
        pvVfsWrite = lSize
        Call CopyMemory(ByVal uFile.BufferPtr, ByVal lPtr, pvVfsWrite)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsWrite
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsWrite = uFile.Stream.VfsWriteFile(lPtr, lSize)
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
    If Not IsEmpty(uFile.BufferArray) Then
        Select Case lMoveMethod
        Case FILE_BEGIN
            pvVfsSeek = lPosition
        Case FILE_CURRENT
            pvVfsSeek = uFile.BufferPtr - uFile.BufferBase + lPosition
        Case FILE_END
            pvVfsSeek = uFile.BufferSize + lPosition
        End Select
        uFile.BufferPtr = uFile.BufferBase + pvVfsSeek
    ElseIf Not uFile.Stream Is Nothing Then
        pvVfsSeek = uFile.Stream.VfsSetFilePointer(lPosition, lMoveMethod)
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
    If Not IsEmpty(uFile.BufferArray) Then
        If uFile.BufferSize <> uFile.BufferPtr - uFile.BufferBase Then
            uFile.BufferSize = uFile.BufferPtr - uFile.BufferBase
            If uFile.BufferSize > 0 Then
                ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            Else
                Erase uFile.BufferArray
            End If
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + uFile.BufferSize
        End If
    ElseIf Not uFile.Stream Is Nothing Then
        uFile.Stream.VfsSetEndOfFile sMetaData
    Else
        If SetEndOfFile(uFile.Handle) = 0 Then
            Err.Raise vbObjectError, , GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Sub

'= common ================================================================

Private Function pvFromOemString(baBuffer() As Byte, Optional ByVal lCodePage As Long = CP_OEMCP) As String
    Dim lSize           As Long
    
    If UBound(baBuffer) >= 0 Then
        lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, 0, 0)
        If lSize > 0 Then
            pvFromOemString = String(lSize, 0)
            lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, ByVal StrPtr(pvFromOemString), lSize)
        End If
    End If
End Function

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

Private Function pvSetError(Optional Source As String, Optional Description As String, Optional ByVal CanContinue As Boolean) As Boolean
    If StrPtr(Source) = 0 And StrPtr(Description) = 0 Then
        m_sLastError = vbNullString
        m_bCancel = False
    Else
        #If ImplLogging Then
            If m_lCurrentFile >= 0 Then
                Debug.Print "Error in " & Source & ": " & Description, m_uFiles(m_lCurrentFile).FileName, Timer
            Else
                Debug.Print "Error in " & Source & ": " & Description, Timer
            End If
        #End If ' ImplLogging
        m_sLastError = Description
        If Not CanContinue Then
            m_bCancel = True
        End If
        RaiseEvent Error(m_lCurrentFile, Source, m_sLastError, m_bCancel)
        pvSetError = m_bCancel
    End If
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
        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    
    If lpThunk = 0 Then
        baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2)
        lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function pvGetCrc32Table() As Long
    Static aTable()     As Long
    Static bIsInit      As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lReminder       As Long
    Dim lValue          As Long

    If Not bIsInit Then
        '--- table mem allocated: 4KB
        ReDim aTable(0 To &H3FF) As Long
        For lIdx = 0 To &H3FF
            If lIdx < &H100 Then
                lReminder = 0
                lValue = lIdx
            Else
                lReminder = aTable(lIdx - &H100)
                lValue = 0
            End If
            For lJdx = 1 To 8
                If ((lReminder Xor lValue) And 1) <> 0 Then
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF Xor &HEDB88320
                Else
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF
                End If
                lValue = (lValue And &HFFFFFFFE) \ 2 And &H7FFFFFFF
            Next
            aTable(lIdx) = lReminder
        Next
        bIsInit = True
    End If
    pvGetCrc32Table = VarPtr(aTable(0))
End Function

'= shared ================================================================

#If ImplUseShared = 0 Then

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
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
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

#If ImplUseApiForBase64 Then

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function

#Else

Private Function FromBase64Array(sText As String) As Byte()
    Const STR_BASE64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim laPow2(0 To 30) As Long
    Dim baInput()       As Byte
    Dim laTable()       As Long
    Dim lIdx            As Long
    Dim lCh             As Long
    Dim baOutput()      As Byte
    Dim lPtr            As Long
    Dim lCarry          As Long
    
    On Error GoTo EH
    '--- init decoding table
    For lIdx = 0 To UBound(laPow2)
        laPow2(lIdx) = 2 ^ lIdx
    Next
    ReDim laTable(0 To &H3FF) As Long
    For lIdx = 0 To UBound(laTable)
        laTable(lIdx) = -1
    Next
    baInput = StrConv(STR_BASE64_CHARS, vbFromUnicode)
    For lIdx = 0 To UBound(baInput)
        lCh = baInput(lIdx)
        laTable(lCh) = lIdx * laPow2(2)
        laTable(&H100 + lCh) = (lIdx And &H30) \ laPow2(4) Or (lIdx And &HF) * laPow2(12)
        laTable(&H200 + lCh) = (lIdx And &H3) * laPow2(22) Or (lIdx And &H3C) * laPow2(6)
        laTable(&H300 + lCh) = lIdx * laPow2(16)
    Next
    '--- figure out output size
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    If laTable(baInput(UBound(baInput))) >= 0 Then
        lIdx = 1
    ElseIf laTable(baInput(UBound(baInput) - 1)) >= 0 Then
        lIdx = 2
    Else
        lIdx = 3
    End If
    ReDim baOutput(0 To (Len(sText) \ 4) * 3 - lIdx) As Byte
    '--- decode loop
    lPtr = VarPtr(baOutput(0))
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lCh = laTable(baInput(lIdx + 0)) _
            Or laTable(&H100 + baInput(lIdx + 1)) _
            Or laTable(&H200 + baInput(lIdx + 2)) _
            Or laTable(&H300 + baInput(lIdx + 3))
        If lCh < 0 Then
            Exit For
        End If
        Call CopyMemory(ByVal lPtr, lCh, 3)
        lPtr = lPtr + 3
    Next
    '--- note: unsigned substract for LARGEADDRESSAWARE processes
    lPtr = ((lPtr Xor &H80000000) - VarPtr(baOutput(0))) Xor &H80000000
    '-- handle trailing 4 chars (or input error)
    If lIdx <= UBound(baInput) Then
        lCh = laTable(baInput(lIdx + 0))
        Debug.Assert lCh >= 0 '--- input error
        If lCh < 0 Then
            GoTo QH
        End If
        lCarry = lCh
        lCh = laTable(baInput(lIdx + 1))
        Debug.Assert lCh >= 0 '--- input error
        If lCh < 0 Then
            GoTo QH
        End If
        baOutput(lPtr + 0) = (lCarry Or lCh \ laPow2(6)) And &HFF
        lCarry = lCh * laPow2(2)
        lCh = laTable(baInput(lIdx + 2))
        If lCh < 0 Then
            lPtr = lPtr + 1
            GoTo QH
        End If
        baOutput(lPtr + 1) = (lCarry Or lCh \ laPow2(4)) And &HFF
        lCarry = lCh * laPow2(4)
        lCh = laTable(baInput(lIdx + 3))
        If lCh < 0 Then
            lPtr = lPtr + 2
            GoTo QH
        End If
        Debug.Assert lCh < 0 '--- error in decoding tables
        baOutput(lPtr + 2) = (lCarry Or lCh \ laPow2(2)) And &HFF
        lPtr = lPtr + 3
    End If
QH:
    '--- will redim only on input error
    If UBound(baOutput) <> lPtr - 1 Then
        ReDim Preserve baOutput(0 To lPtr - 1) As Byte
    End If
    FromBase64Array = baOutput
    Exit Function
EH:
    Debug.Print Err.Description
End Function

#End If

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

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
    m_lCurrentFile = -1
End Sub
