# THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THIS CODE AND INFORMATION ARE NOT SUPPORTED BY XEBIALABS. 
import sys

from java.lang import Exception
from java.io import PrintWriter
from java.io import StringWriter

from com.xebialabs.overthere import CmdLine, ConnectionOptions, OperatingSystemFamily, Overthere
from com.xebialabs.overthere.cifs import CifsConnectionBuilder
from com.xebialabs.overthere.cifs import CifsConnectionType
from com.xebialabs.overthere.util import CapturingOverthereExecutionOutputHandler, OverthereUtils

class WinrmRemoteCScript():
    def __init__(self, username, password, address, connectionType, timeout, allowDelegate, cscriptExecutable, remotePath, script):
        self.options = ConnectionOptions()
        self.options.set(ConnectionOptions.USERNAME, username)
        self.options.set(ConnectionOptions.PASSWORD, password)
        self.options.set(ConnectionOptions.ADDRESS, address)
        self.options.set(ConnectionOptions.OPERATING_SYSTEM, OperatingSystemFamily.WINDOWS)

        self.remotePath = remotePath
        self.script = script
        self.connectionType = connectionType
        # WINRM_NATIVE only
        self.allowDelegate = allowDelegate
        # WINRM_INTERNAL only
        self.timeout = timeout

        self.stdout = CapturingOverthereExecutionOutputHandler.capturingHandler()
        self.stderr = CapturingOverthereExecutionOutputHandler.capturingHandler()

    def customize(self, options):
        if self.connectionType == 'WINRM_NATIVE':
            options.set(CifsConnectionBuilder.CONNECTION_TYPE, CifsConnectionType.WINRM_NATIVE)
            options.set(CifsConnectionBuilder.WINRS_ALLOW_DELEGATE, allowDelegate)
        elif self.connectionType == 'WINRM_INTERNAL':
            options.set(CifsConnectionBuilder.CONNECTION_TYPE, CifsConnectionType.WINRM_INTERNAL)
            options.set(CifsConnectionBuilder.WINRM_KERBEROS_USE_HTTP_SPN, True)
            options.set(CifsConnectionBuilder.WINRM_TIMEMOUT, timeout);
        #print 'DEBUG: Options:', options

    def execute(self):
        self.customize(self.options)
        connection = None
        try:
            connection = Overthere.getConnection(CifsConnectionBuilder.CIFS_PROTOCOL, self.options)
            connection.setWorkingDirectory(connection.getFile(self.remotePath))
            # upload the script and pass it to cscript.exe
            targetFile = connection.getTempFile('uploaded-script', '.vbs')
            OverthereUtils.write(String(self.script).getBytes(), targetFile)
            targetFile.setExecutable(True)
            # run cscript in batch mode
            scriptCommand = CmdLine.build(cscriptExecutable, '//B', '//nologo', targetFile.getPath())
            return connection.execute(self.stdout, self.stderr, scriptCommand)
        except Exception, e:
            stacktrace = StringWriter()
            writer = PrintWriter(stacktrace, True)
            e.printStackTrace(writer)
            self.stderr.handleLine(stacktrace.toString())
            return 1
        finally:
            if connection is not None:
                connection.close()

    def getStdout(self):
        return self.stdout.getOutput()

    def getStdoutLines(self):
        return self.stdout.getOutputLines()

    def getStderr(self):
        return self.stderr.getOutput()

    def getStderrLines(self):
        return self.stderr.getOutputLines()

qtpScript = """
Dim qtApp
Dim qtTest
Set qtApp = CreateObject("QuickTest.Application") 
If Not qtApp.launched Then 
  qtApp.Launch 
End If 

qtApp.Visible = False
qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Fast"
qtApp.Options.Run.ViewResults = False
qtApp.Open "%s", True 

Set qtTest = qtApp.Test
qtTest.Settings.Run.OnError = "NextStep" 
qtTest.Run

'print even when run via cscript //B
Set fso = CreateObject("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream(1)
stdout.WriteLine qtTest.LastRunResults.Status
Set stdout = Nothing
Set fso = Nothing

qtTest.Close 
qtApp.quit

Set qtTest = Nothing
Set qtApp = Nothing
""" % (testPath)

script = WinrmRemoteCScript(username, password, address, connectionType, timeout, allowDelegate, cscriptExecutable, remotePath, qtpScript)
exitCode = script.execute()

output = script.getStdout()
err = script.getStderr()

if (exitCode == 0):
    print output
else:
    print "Exit code "
    print exitCode
    print
    print "#### Output:"
    print output

    print "#### Error stream:"
    print err
    print
    print "----"

    sys.exit(exitCode)
