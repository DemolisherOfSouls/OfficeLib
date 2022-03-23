Attribute VB_Name = "GlobalVars"
Option Explicit
Option Compare Text
Option Base 1

'Quality Program Global Variables

Public Const CSEpicor As String = "Provider=SQLOLEDB;Data Source=esql.ashworth.com;Initial Catalog=Ashworth;User ID=devapp;Password=d3v@PP;"
Public Const CSEngineer As String = "Provider=SQLOLEDB;Data Source=esql.ashworth.com;Initial Catalog=Engineering;User ID=Engineering;Password=Engineer2015;"

Public Const MBDataMissingContact As String = "Data Required For This Inspection is Missing Please Contact Continuous Improvement or Engineering"
Public Const MBDataErrorsResbmit As String = "Errors have been Detected in the Data Entered Please Correct and Resubmit"
Public Const MBExitDisabled As String = "The X is disabled, please use a button on the form"
Public Const MBFillOutResubmit As String = "Please Completely Fill out the Form and Resubmit"
Public Const MBErrorOpComments As String = "Error Returning the Operation Comments"
Public Const MBFillOutSpiral As String = "Please Completely Fill out the Form and Resubmit. The Spiral Hand and Machine Number Must Be Entered"
Public Const MBNoData As String = "No Data is Available"

Public DBEpicor As New ADODB.Connection
Public DBEng As New ADODB.Connection

Public JobNum As String
Public Inspection As String
Public Operation As String
Public DiffSpiralCount As Integer
Public CommentBoxReturn As String
Public BeltType As String
Public Company As Integer
Public SampleNum As Integer
Public Employee As String

Public BeltWidth As Variant
Public CrimpDepth As Variant
Public SpiralSize As String
Public Belt_Category As String
Public Loops As Integer
Public Center_Link_Location As Variant
Public Rod_Diam As Variant
Public Fabric_Width As Variant
