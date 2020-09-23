<div align="center">

## Transfering Files With The Internet Transfer Control


</div>

### Description

Thanks to the Internet's ever-increasing prominence in our world, we developers are constantly finding new and better ways to take advantage of its capabilities. Frequently, that means finding new ways to perform tasks on the Internet - pushing the limits to do something that hasn't been done before. At other times, we must find alternate paths to take advantage of functionality that's existed for years, such as file transfers using the File Transfer Protocol (FTP). FTP gives us the ability to send or receive all sorts of files across the Internet. Web browsers use underlying FTP functionality when downloading files. We can employ that same functionality in our Visual Basic applications to transfer files across the Internet or intranet by using the Microsoft Internet Transfer Control.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max \- Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-demian-net.md)
**Level**          |Advanced
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-demian-net-transfering-files-with-the-internet-transfer-control__1-8841/archive/master.zip)





### Source Code

<center>
<img src="http://demiannet.hypermart.net/artic1.jpg" border="0">
</center>
<p class=title>A Look At The Sample Project</p><p> </p><p><b>Figure A:</b> Our sample application looks like this at design time.<br></p>
  <img alt="Figure A" border="0" src="http://www.dev-center.com/data/article_images/000490_1.gif">
  <p> </p>
  <p>We'll also use the status bar control that ships with VB to display
  status information about the connection. To add the status-bar control to
  your application, choose the Project | Components... menu item, then select
  Microsoft Windows Common Controls 5.0 and click OK.</p>
  <p> </p>
  <p>The sample application lets users transfer files by dragging them from
  one list box to the other. Although this feature is optional, it's very
  user-friendly. The sample code found in Listing A shows you the steps to
  implement this drag-and-drop feature--pay particular attention to the
  MouseDown, MouseUp, and DragDrop events. We'll examine how those transfers
  are performed throughout the rest of this article.</p>
  <p> </p>
<p class=title>Using The Control</p><p> </p><p>The Internet Transfer Control provides fairly extensive capabilities for
  transferring data across the Internet, in the form of either Web pages or
  files. For our purposes, we'll concentrate on file transfers and leave the
  rest for another article. The control resides in the MSINET.OCX file. To
  load the control into your VB toolbox, choose Project | Components.... Next,
  find the Microsoft Internet Transfer Control 5.0 control, select it by
  placing an X beside it, then click OK. Now, add the control to your project
  form. Note that the control will appear as a button and won't be visible at
  runtime.</p>
  <p> </p>
  <p>You can open the Object Browser (by pressing [F2]) to examine all the
  properties, methods, events, and built-in constants available through this
  code. In addition to the control's help file, this information makes an
  excellent reference. For this article, we'll focus on the small set of
  available properties and methods listed in Table A.
  <p> </p>
  <p><b>Table A: </b>Selected properties, methods, and events
  <table border="0">
   <tbody>
    <tr vAlign="top">
     <td align="left"><b>Properties</b></td>
     <td align="left"><b>Description</b></td>
    </tr>
    <tr vAlign="top">
     <td align="left">Password</td>
     <td align="left">The password you use when connecting with the FTP
      server.</td>
    </tr>
    <tr vAlign="top">
     <td align="left">StillExecuting</td>
     <td align="left">Specifies whether a command is still being processed.</td>
    </tr>
    <tr vAlign="top">
     <td align="left">URL</td>
     <td align="left">The URL of the FTP server.</td>
    </tr>
    <tr vAlign="top">
     <td align="left">Username</td>
     <td align="left">User name to use to log into the FTP server.</td>
    </tr>
    <tr vAlign="top">
     <td align="left"><b>Methods</b></td>
     <td align="left"><b>Description</b></td>
    </tr>
    <tr vAlign="top">
     <td align="left">Execute</td>
     <td align="left">Initiates an asynchronous command/connection.</td>
    </tr>
    <tr vAlign="top">
     <td align="left">GetChunk</td>
     <td align="left">Reads data from the buffer.</td>
    </tr>
    <tr vAlign="top">
     <td align="left">OpenURL</td>
     <td align="left">Initiates a synchronous command/connection.</td>
    </tr>
    <tr vAlign="top">
     <td align="left"><b>Events</b></td>
     <td align="left"><b>Description</b></td>
    </tr>
    <tr vAlign="top">
     <td align="left">StateChanged</td>
     <td align="left">Fires whenever the control state has changed, for
      example, when a response is received from the FTP server.</td>
    </tr>
   </tbody>
  </table>
  <p>
<p class=title>Performing Transfers</p><p> </p><p>To perform FTP transfers, you must follow a few basic steps. First, you
  define the FTP server you want to attach to. You can specify the FTP site in
  two ways: using the <b>RemoteHost</b> and <b>RemotePort</b> properties or
  via the <b>URL</b> property. For simplicity's sake, we'll use the <b>URL</b>
  property:</p>
  <p> </p>
  <p>Inet1.URL = txtURL<br>
  </p>
  You also must specify the user name and password you'll provide. Many FTP
  sites allow anonymous connections. In those cases, the user name <i>anonymous</i>
  will work with any password you like, although most FTP sites ask you to
  provide your E-mail address, as well. Here's the syntax:
  <p>
  <p>Inet1.UserName = txtUsername<br>
  Inet1.Password = txtPassword<br>
  </p>
  Setting the <b>URL</b> property will clear the <b>Username</b> and <b>Password</b>
  properties. So, be sure to set the URL first, then specify the user name and
  password.
  <p>
  <p>Since we're going to be dealing strictly with FTP connections, we'll set
  the <b>Protocol</b> property accordingly, as follows:</p>
  <p>
  <p>Inet1.Protocol = icFTP</p>
  <p> </p>
  We'll want to execute these commands when we make our first connection to
  the FTP server. We use the cmdConnect command button to establish this
  connection, so the code will go to the server. At the same time, when we
  make this first FTP connection, we'll also retrieve the list of files
  available on the FTP server. We'll see how to do this next.
  <p>
<p class=title>Executing Gets Things Done</p><p> </p><p>You'll use the Execute method to send all commands to the FTP site through
  the control. The syntax of the Execute method is</p>
  <p>
  <p>Inet1.Execute URL, Operation, Data, _<br>
    RequestHeaders<br>
</p>
  However, when performing FTP commands, we'll only use the URL and Operation
  parameters. The others have no meaning for us--they're used in other
  processes. You send all FTP commands in the Operation parameter; they take
  the syntax command [file1 [file2]]. The help file for the Internet Transfer
  Control includes a list of valid FTP commands under the Execute page. We'll
  focus on a few of these commands in the rest of this article.
<p class=title>Asynchronous Processing</p><p> </p>When you're using the Execute method, keep in mind that all its operations
  are <i>asynchronous</i>. This means that when you tell the control to
  perform an operation, it starts the operation but returns control back to
  the application. The control will handle all communications back and forth,
  based on the properties and commands you've given it. When the operation is
  completed, the control will notify the application. If you use the OpenURL
  method, the control makes a <i>synchronous</i> connection and executes the
  command. However, control doesn't return until the command finishes
  executing. This more straightforward approach is somewhat simpler to
  program. Since the asynchronous approach is more flexible--and therefore
  preferable--we'll use it exclusively here.
  <p>Our discussion of the asynchronous approach would be incomplete without
  mentioning the <b>StillExecuting</b> property. This property identifies when
  the control is in the middle of performing some operation. If you need to
  perform an operation that requires several commands, you'll start the first
  command, loop until the control has stopped processing the command (i.e., <b>StillExecuting</b>
  is False), then move on to the next operation, as follows:
  <p>
  <p>Inet1.Execute txtURL, "get MyFile.txt"<br>
  Do <br>
    DoEvents<br>
  Loop While Inet1.StillExecuting<br>
  </p>
  In event-driven programming, we want to be able to react when the operation
  is complete. We'll use the StateChanged event to provide this functionality.
  Specifically, we'll look for the new State of the control to be
  icResponseCompleted. It may be useful to set a variable, such as iLastFTP,
  to store a value signifying which FTP command executed last. Then you can
  test that variable in the StateChanged event to determine what command
  completion you're reacting to, with the lines:
  <p>
  <p>Sub Inet1.StateChanged(ByVal State As Integer)<br>
   Select Case State<br>
    Case icResponseCompleted<br>
      `put your code here<br>
   End Select<br>
  End Sub<br>
  </p>
  Of course, the State parameter can hold a number of other values as well. We
  show these in the full code listing, found in Listing A. You can also check
  the help file for all these values. Now, let's build our FTP application.
<p class=title>Creating The Sample Project</p><p> </p>The first step is to begin a new EXE project in VB5. Build a form similar to
  that shown in Figure A. As you can see, the form should include TextBox
  controls for the target URL, user name, and password. You'll also need to
  provide a way to display both local and remote files. Our example uses the
  DirListBox, DriveListBox, and FileListBox controls for the local files, and
  a standard ListBox control to display the remote files. Finally, you must
  add a CommandButton to establish the initial connection. After that, our
  work with the list boxes will be complete. Table B shows the controls to add
  to the form, as well as some key properties.
  <p> </p>
  <p><b>Table B:</b> Controls to add to the form
  <table border="0">
   <tbody>
    <tr vAlign="bottom">
     <td align="left"><b>Control</b></td>
     <td align="left"><b>Property</b></td>
     <td align="left"><b>Setting</b></td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">Form</td>
     <td align="left">Caption</td>
     <td align="left">File Transfer</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">TextBox</td>
     <td align="left">Name</td>
     <td align="left">txtURL</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">TextBox</td>
     <td align="left">Name</td>
     <td align="left">txtUserName</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">TextBox</td>
     <td align="left">Name</td>
     <td align="left">txtPassword</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">DriveListBox</td>
     <td align="left">Name</td>
     <td align="left">drvLocal</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">DirListBox</td>
     <td align="left">Name</td>
     <td align="left">dirLocal</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">FileListBox</td>
     <td align="left">Name</td>
     <td align="left">filLocal</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">ListBox</td>
     <td align="left">Name</td>
     <td align="left">lstRemoteFiles</td>
    </tr>
    <tr vAlign="bottom">
     <td align="left">CommandButton</td>
     <td align="left">Name</td>
     <td align="left">cmdConnect</td>
    </tr>
    <tr vAlign="bottom">
     <td align="right"></td>
     <td align="left">Caption</td>
     <td align="left">Connect</td>
    </tr>
   </tbody>
  </table>
  <p>
  <p>As the name implies, the Connect button will connect to the designated
  FTP site and retrieve a list of files. To accomplish this, put the following
  code in the cmdConnect_Click event:</p>
  <p>
  <p>Public Sub cmdConnect_Click()<br>
    Inet1.URL = txtURL<br>
    Inet1.UserName = txtUserName<br>
    Inet1.Password = txtPassword<br>
    Inet1.Protocol = icFTP<br>
    `Use constant to identify that<br>
    `we're getting a directory listing.<br>
    `We'll use it in the<br>
    `Inet1_StateChanged Event.<br>
    iLastFTP = ftpDIR<br>
    Inet1.Execute Inet1.URL, "DIR"<br>
  End Sub<br>
  </p>
  The first time we use the Execute method, we establish a connection between
  the user machine and the FTP site. Executing the Dir command will place a
  list of files in the control's buffer. To retrieve them from the buffer,
  we'll use the GetChunk method. GetChunk requires one parameter that
  specifies the maximum amount of data (in bytes) that we'll retrieve. We
  specify an amount and keep looping until we've emptied out the buffer. The
  result is a string of filenames, separated by a carriage return and line
  feed (vbCrLf). We can then display the list of files however we want. We
  wrote the function ShowRemoteFileList() in Listing A to load the file list
  into a list box.
  <p> </p>
  <p>Once we have the list of files, we can let the user upload files to (or
  download files from) the FTP site. Since downloading files is more common,
  let's consider this example first. We download files by executing the FTP
  command Get. The syntax of the command is Get file1 file2, where <i>file1</i>
  is the name of the file on the FTP site and <i>file2</i> is the name you
  want the file to have locally. File2 can include path information as well.
  The GetFiles() function in Listing A demonstrates how to issue the command
  and retrieve the file.
  <p> </p>
  <p>Similarly, if you want to upload a file to the FTP site (assuming you
  have write privileges at the site), you use the FTP command Put. The syntax
  of the command is Put file1 file2, where <i>file1</i> is the local filename
  (which can include the path) and <i>file2</i> is the name the file will have
  on the FTP site. The PutFiles() function in Listing A demonstrates this
  process. Please note that you'll have a problem to work around. The FTP
  Command Line doesn't allow spaces in the filename or path. To solve this
  problem, you can take one of the following steps:
  <p> </p>
  <p>1. Use relative paths when specifying local files (which is the option we
  used in the sample program).
  <p>2. Place quotation marks (Chr(34)) around the full path and filename
  (such as <i>C:\My FTP Files\TestFile.txt</i>) in the ftp command.
  <p>3. Use the 8.3-character directory name
  <p>4. Don't allow spaces in directory names.
  <p> </p>
  <p>With a little work, our application can allow the user to select multiple
  files to transfer in one operation. Of course, the application will need to
  issue an Execute command for each transfer. Then, we must test the <b>StillExecuting</b>
  property to determine whether the control has finished executing that
  command. Once it's complete, we can loop back and send the command again for
  the second file. We can continue this process for as many files as
  necessary.</p>
<p class=title>Known Bugs And Issues</p><p> </p>You should be aware of several issues that exist with the current versions
  of the control. These issues vary depending on which version you're using.
  In the version that ships with VB 5 (version 5.00.3714), the control sends
  all filenames as uppercase when you're sending or receiving files. If you're
  hitting an Internet Information Server (IIS) using NT/DOS file settings,
  case doesn't matter, since the filenames aren't case-sensitive. However, if
  you're hitting a UNIX server, it's extremely important, since UNIX filenames
  <i>are</i> case-sensitive. The result is that any files you send will be
  named in all uppercase, and you won't be able to retrieve files that have
  lowercase letters in their names.
  <p>
  <p>Fortunately, Microsoft is aware of this conflict (see the Microsoft
  Knowledge Base article <a href="http://support.microsoft.com/support/kb/articles/Q168/7/66.asp" target="_blank">support.microsoft.com/support/kb/articles/Q168/7/66.asp</a>
  for more information) and has corrected it in Service Pack 2 for Visual
  Studio. However, the SP2 control (version 5.01.4319) introduces an even
  worse problem.</p>
  <p> </p>
  <p>In the SP2 version of the control, you can't log in to any server, other
  than a strictly anonymous server (such as <a href="ftp://ftp.microsoft.com" target="_blank">ftp://ftp.microsoft.com</a>).
  User names and passwords are sent incorrectly to the FTP server. (See the
  Microsoft Knowledge Base article <a href="http://support.microsoft.com/support/kb/articles/Q173/2/65.asp" target="_blank">support.microsoft.com/support/kb/articles/Q173/2/65.asp</a>
  for more details.)
  <p> </p>
  <p>Finally, Microsoft released Service Pack 3 (<a href="http://msdn.microsoft.com/vstudio/sp/vs6sp3/default.asp" target="_blank">http://msdn.microsoft.com/vstudio/sp/vs6sp3/default.asp</a>)
  in early December 1997, correcting these problems.</p>
<p class=title>Code For Core Functionality</p><p> </p><p>Add the following to a form:
  <p> </p>
  <p><font face="Courier New">Private Const ftpDIR As Integer = 0<br>
  Private Const ftpPUT As Integer = 1<br>
  Private Const ftpGET As Integer = 2<br>
  Private Const ftpDEL As Integer = 3<br>
  Private iLastFTP As Integer<br>
  </font></p>
  <p><font face="Courier New">Private Sub cmdConnect_Click()<br>
    On Error GoTo ConnectError<br>
    Inet1.URL = txtURL<br>
    Inet1.UserName = txtUserName<br>
    Inet1.Password = txtPassword<br>
    Inet1.Protocol = icFTP<br>
    iLastFTP = ftpDIR<br>
  <br>
    Inet1.Execute Inet1.URL, "DIR"<br>
  End Sub<br>
  <br>
  Private Sub Inet1_StateChanged(ByVal _<br>
    State As Integer)<br>
    Select Case State<br>
      Case icNone<br>
  sbFTP.Panels("status").Text = ""<br>
      Case icResolvingHost<br>
  sbFTP.Panels("status").Text<br>
        =
  "Resolving Host"<br>
      Case icHostResolved<br>
  sbFTP.Panels("status").Text _<br>
        = "Host
  Resolved"<br>
      Case icConnecting<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Connecting..."<br>
      Case icConnected<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Connected!"<br>
      Case icRequesting<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Requesting..."<br>
      Case icRequestSent<br>
  sbFTP.Panels("status").Text _<br>
        = "Request
  Sent"<br>
      Case icReceivingResponse<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Receiving Response..."<br>
      Case icResponseReceived<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Response Received!"<br>
      Case icDisconnecting<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Disconnecting..."<br>
  <br>
      Case icDisconnected<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Disconnected"<br>
      Case icError<br>
  sbFTP.Panels("status").Text _<br>
        = "Error!
  " & Trim(CStr( _<br>
  Inet1.ResponseCode)) & _<br>
        ": "
  & Inet1.ResponseInfo<br>
      Case icResponseCompleted<br>
  sbFTP.Panels("status").Text _<br>
        =
  "Response Completed!"<br>
  ReactToResponse iLastFTP<br>
    End Select<br>
  End Sub<br>
  <br>
  Public Function _<br>
    ReactToResponse(ByVal _<br>
    iLastCommand As Integer) As Long<br>
    Select Case iLastCommand<br>
      Case ftpDIR<br>
  ShowRemoteFileList<br>
      Case ftpPUT<br>
        MsgBox
  "File Sent from " & CurDir()<br>
      Case ftpGET<br>
        MsgBox
  "File Received "& "in " & CurDir()<br>
      Case ftpDEL<br>
    End Select<br>
  End Function<br>
  <br>
  Public Function ShowRemoteFileList() As Long<br>
    Dim sFileList As String<br>
    Dim sTemp As String<br>
    Dim p As Integer<br>
    sTemp = Inet1.GetChunk(1024)<br>
    Do While Len(sTemp) > 0<br>
      DoEvents<br>
      sFileList = sFileList & sTemp<br>
      sTemp = Inet1.GetChunk(1024)<br>
    Loop<br>
    lstRemoteFiles.Clear<br>
    Do While sFileList > ""<br>
      DoEvents<br>
      p = InStr(sFileList, vbCrLf)<br>
      If p > 0 Then<br>
  lstRemoteFiles.AddItem <br>
  Left(sFileList, p - 1)<br>
        If
  Len(sFileList) > (p + 2) Then<br>
  sFileList = Mid(sFileList, p + 2)<br>
        Else<br>
  sFileList = ""<br>
        End If<br>
      Else<br>
  lstRemoteFiles.AddItem sFileList<br>
        sFileList
  = ""<br>
      End If<br>
    Loop<br>
  End Function<br>
  </font></p><p class=title>Code For Core Functionality Part 2</p><p> </p><p>'Continued:</p>
  <p><font face="Courier New">Public Function GetFiles(sFileList As String) As
  Long<br>
    Dim sFile As String<br>
    Dim sTemp As String<br>
    Dim p As Integer<br>
    iLastFTP = ftpGET<br>
    sTemp = sFileList<br>
    Do While sTemp > ""<br>
      DoEvents<br>
      p = InStr(sTemp, "|")<br>
      If p Then<br>
        sFile =
  Left(sTemp, p - 1)<br>
        sTemp =
  Mid(sTemp, p + 1)<br>
      Else<br>
        sFile =
  sTemp<br>
        sTemp =
  ""<br>
      End If<br>
      Inet1.Execute Inet1.URL,
  "GET " & sFile & _<br>
        "
  " & sFile<br>
    'wait until this execution is done <br>
    `before going to next file<br>
      Do<br>
        DoEvents<br>
      Loop Until Not _<br>
  Inet1.StillExecuting<br>
    Loop<br>
    iLastFTP = ftpDIR<br>
    Inet1.Execute Inet1.URL, "DIR"<br>
  End Function<br>
  </font></p>
  <p><font face="Courier New">Public Function PutFiles(sFileList As String) As
  Long<br>
    Dim sFile As String<br>
    Dim sTemp As String<br>
    Dim p As Integer<br>
    iLastFTP = ftpPUT<br>
    sTemp = sFileList<br>
    Do While sTemp > ""<br>
      DoEvents<br>
      p = InStr(sTemp, "|")<br>
      If p Then<br>
        sFile =
  Left(sTemp, p - 1)<br>
        sTemp =
  Mid(sTemp, p + 1)<br>
      Else<br>
        sFile =
  sTemp<br>
        sTemp =
  ""<br>
      End If<br>
      Inet1.Execute Inet1.URL,
  "PUT" & sFile & _<br>
        "
  " & sFile<br>
    'wait until this execution is done <br>
    `before going to next file<br>
      Do<br>
        DoEvents<br>
      Loop Until Not
  Inet1.StillExecuting<br>
    Loop<br>
    iLastFTP = ftpDIR<br>
    Inet1.Execute Inet1.URL, "DIR"<br>
  End Function<br>
  <br>
  Private Sub dirLocal_Change()<br>
    filLocal.Path = dirLocal.Path<br>
  End Sub<br>
  <br>
  Private Sub drvLocal_Change()<br>
    dirLocal.Path = drvLocal.Drive<br>
  End Sub<br>
  <br>
  Private Sub filLocal_DragDrop(Source _<br>
      As Control, X As Single, Y As
  Single)<br>
    'receiving files from FTP site.<br>
    Dim I As Integer<br>
    Dim sFileList As String<br>
    If TypeOf Source Is ListBox Then<br>
      For i = 0 _<br>
        To
  Source.ListCount - 1<br>
        If
  Source.Selected(i) Then<br>
  sFileList = _<br>
  sFileList & _<br>
  Source.List(i) & "|"<br>
        End If<br>
      Next<br>
    End If<br>
    If Len(sFileList) > 0 Then<br>
      'strip off the last pipe<br>
      sFileList = Left(sFileList, _<br>
  Len(sFileList) - 1)<br>
      GetFiles sFileList<br>
    End If<br>
  End Sub<br>
  <br>
  Private Sub _<br>
    filLocal_MouseDown(Button As _<br>
    Integer, Shift As Integer, X As _<br>
    Single, Y As Single)<br>
    filLocal.Drag vbBeginDrag<br>
  End Sub<br>
  <br>
  Private Sub filLocal_MouseUp(Button _<br>
    As Integer, Shift As Integer, _<br>
    X As Single, Y As Single)<br>
    filLocal.Drag vbEndDrag<br>
  End Sub<br>
  </font></p>
<p class=title>Code For Core Functionality Part 3</p><p> </p><p>'Continued:</p>
  <p> </p>
  <p><font face="Courier New">Private Sub _<br>
    lstRemoteFiles_DragDrop(Source _<br>
    As Control, X As Single, Y As Single)<br>
    Dim I As Integer<br>
    Dim sFileList As String<br>
    If TypeOf Source Is FileListBox Then<br>
      For i = 0 To Source.ListCount - 1<br>
        If
  Source.Selected(i) Then<br>
  sFileList = sFileList & _<br>
  Source.List(i) & "|"<br>
        End If<br>
      Next<br>
    End If<br>
    If Len(sFileList) > 0 Then<br>
      'strip off the last pipe<br>
      sFileList = Left(sFileList, _<br>
  Len(sFileList) - 1)<br>
      PutFiles sFileList<br>
    End If<br>
  End Sub<br>
  <br>
  Private Sub _<br>
    lstRemoteFiles_KeyDown(KeyCode _<br>
    As Integer, Shift As Integer)<br>
    If KeyCode = vbKeyDelete Then<br>
      Inet1.Execute Inet1.URL,
  "DEL " & _<br>
  lstRemoteFiles.List( _<br>
  lstRemoteFiles.ListIndex)<br>
      Do<br>
        DoEvents<br>
      Loop While Inet1.StillExecuting<br>
    End If<br>
    iLastFTP = ftpDIR<br>
    Inet1.Execute Inet1.URL, "DIR"<br>
  End Sub<br>
  <br>
  Private Sub _<br>
    lstRemoteFiles_MouseDown(Button _<br>
    As Integer, Shift As Integer, )<br>
    X As Single, Y As Single)<br>
    lstRemoteFiles.Drag vbBeginDrag<br>
  End Sub<br>
  <br>
  Private Sub lstRemoteFiles_MouseUp(Button As _<br>
    Integer, Shift As Integer, _<br>
    X As Single, Y As Single)<br>
    lstRemoteFiles.Drag vbEndDrag<br>
  End Sub</font></p>
<p class=title>Conclusion</p><p> </p><p>As the Internet's importance grows in our daily lives, we must make our
  applications more Internet-aware. Actually, the Internet offers several
  solutions to some potential problems--the challenge is to take advantage of
  the existing capabilities to meet those challenges. If you need to transfer
  files between two Internet sites, the Internet Transfer Control offers a
  quick solution. In this article, we've shown you how to use the control in
  your applications. We've also pointed out a couple of bugs to work around.</p>

