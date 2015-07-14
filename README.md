# xlr-qtp-plugin

### Introduction

![image](documentation/qtp-icon.png) 

QuickTest Professional (QTP) is now known as HP Unified Functional Testing.  This plugin should work with either product to support the execution of test sets via the products' Visual Basic Script (VBS) interface.

### Supported Tasks

#### Remote CScript Task

Invokes a VBS script on a Windows machine via WinRM, e.g. to execute a QTP test set. An example script that might be used for QTP is:

```
Dim qtApp
Dim qtTest

'Create the QTP Application object
Set qtApp = CreateObject("QuickTest.Application") 

'If QTP is notopen then open it
If  qtApp.launched <> True then 

qtApp.Launch 

End If 

'Make the QuickTest application visible
qtApp.Visible = True

'Set QuickTest run options
'Instruct QuickTest to perform next step when error occurs

qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Fast"
qtApp.Options.Run.ViewResults = False

'Open the test in read-only mode
qtApp.Open "C:\Program Files\HP\QuickTest Professional\Tests\trial", True 

'set run settings for the test
Set qtTest = qtApp.Test

'Instruct QuickTest to perform next step when error occurs
qtTest.Settings.Run.OnError = "NextStep" 

'Run the test
qtTest.Run

'Check the results of the test run
MsgBox qtTest.LastRunResults.Status

' Close the test
qtTest.Close 

'Close QTP
qtApp.quit

'Release Object
Set qtTest = Nothing
Set qtApp = Nothing
```

The task returns the standard out, standard error and exit code from the script.  You can use this information later in your release template to determine the test result and make decisions during the release.

![image](documentation/QTP_Step.png)

**Input properties**

The task provides the standard properties of the [Windows remote script task](https://docs.xebialabs.com/xl-release/concept/introduction-to-the-xl-release-remote-script-plugin.html).