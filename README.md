<div align="center">

## MSDN WinSock Tutorial


</div>

### Description

This Article teaches the basics to WinSock. I did not write any part of it, and give all credit to Microsoft. I got it straight from the help files in VB, but I am submitting it to get the information out for the selected few that don't already know about WinSock!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Zaczkowski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-zaczkowski.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-zaczkowski-msdn-winsock-tutorial__1-37234/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Internet Assistant for Word Version 3.0">
</HEAD>
<BODY>
<B><U><FONT SIZE=6><P ALIGN="CENTER">Using the Winsock Control</P>
</U><P ALIGN="CENTER"></P>
</B></FONT><P>A WinSock control allows you to connect to a remote machine and exchange data using either the User Datagram Protocol (UDP) or the Transmission Control Protocol (TCP). Both protocols can be used to create client and server applications. Like the Timer control, the WinSock control doesn't have a visible interface at run time.</P>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>Possible Uses</P><DIR>
<DIR>
</FONT><FONT FACE="Symbol" SIZE=5><P>·&#9;</B></FONT>Create a client application that collects user information before sending it to a central server.<BR>
</P>
<FONT FACE="Symbol"><P>·&#9;</FONT>Create a server application that functions as a central collection point for data from several users.<BR>
</P>
<FONT FACE="Symbol"><P>·&#9;</FONT>Create a "chat" application. </P></DIR>
</DIR>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>Selecting a Protocol</P>
</B></FONT><P>When using the WinSock control, the first consideration is whether to use the TCP or the UDP protocol. The major difference between the two lies in their connection state: </P><DIR>
<DIR>
<FONT FACE="Symbol"><P>·&#9;</FONT>The TCP protocol control is a connection-based protocol, and is analogous to a telephone - the user must establish a connection before proceeding.<BR>
</P>
<FONT FACE="Symbol"><P>·&#9;</FONT>The UDP protocol is a connectionless protocol, and the transaction between two computers is like passing a note: a message is sent from one computer to another, but there is no explicit connection between the two. Additionally, the maximum data size of individual sends is determined by the network. </P></DIR>
</DIR>
<P>The nature of the application you are creating will generally determine which protocol you select. Here are a few questions that may help you select the appropriate protocol: </P><DIR>
<DIR>
<P>Will the application require acknowledgment from the server or client when data is sent or received? If so, the TCP protocol requires an explicit connection before sending or receiving data.<BR>
</P>
<P>Will the data be extremely large (such as image or sound files)? Once a connection has been made, the TCP protocol maintains the connection and ensures the integrity of the data. This connection, however, uses more computing resources, making it more "expensive."<BR>
</P>
<P>Will the data be sent intermittently, or in one session? For example, if you are creating an application that notifies specific computers when certain tasks have completed, the UDP protocol may be more appropriate. The UDP protocol is also more suited for sending small amounts of data. </P></DIR>
</DIR>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>Setting the Protocol</P>
</B></FONT><P>To set the protocol that your application will use: at design-time, on the Properties window, click Protocol and select either sckTCPProtocol, or sckUDPProtocol. You can also set the Protocol property in code, as shown below:</P>
<FONT FACE="Courier New" SIZE=2><P>Winsock1.Protocol = sckTCPProtocol</P>
</FONT><B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>Determining the Name of Your Computer</P>
</B></FONT><P>To connect to a remote computer, you must know either its IP address or its "friendly name." The IP address is a series of three digit numbers separated by periods (xxx.xxx.xxx.xxx). In general, it's much easier to remember the friendly name of a computer. </P>
<B><P>To find your computer's name</B> </P><DIR>
<DIR>
<P>On the <B>Taskbar</B> of your computer, click <B>Start</B>.<BR>
</P>
<P>On the <B>Settings</B> item, click the <B>Control Panel</B>.<BR>
</P>
<P>Double-click the <B>Network</B> icon.<BR>
</P>
<P>Click the <B>Identification </B>tab.<BR>
</P>
<P>The name of your computer will be found in the <B>Computer name</B> box. </P></DIR>
</DIR>
<P>Once you have found your computer's name, it can be used as a value for the RemoteHost property.</P>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>TCP Connection Basics</P>
</B></FONT><P>When creating an application that uses the TCP protocol, you must first decide if your application will be a server or a client. Creating a server means that your application will "listen," on a designated port. When the client makes a connection request, the server can then accept the request and thereby complete the connection. Once the connection is complete, the client and server can freely communicate with each other.</P>
<P>The following steps create a rudimentary server:</P>
<B><P>To create a TCP server</B> </P><DIR>
<DIR>
<P>Create a new Standard EXE project.<BR>
</P>
<P>Change the name of the default form to frmServer.<BR>
</P>
<P>Change the caption of the form to "TCP Server."<BR>
</P>
<P>Draw a Winsock control on the form and change its name to tcpServer.<BR>
</P>
<P>Add two TextBox controls to the form. Name the first txtSendData, and the second txtOutput.<BR>
</P>
<P>Add the code below to the form. </P>
<FONT FACE="Courier New" SIZE=2><P>Private Sub Form_Load()</P>
<P>  ' Set the LocalPort property to an integer.</P>
<P>  ' Then invoke the Listen method.</P>
<P>  tcpServer.LocalPort = 1001</P>
<P>  tcpServer.Listen </P>
<P>  frmClient.Show ' Show the client form.</P>
<P>End Sub</P>
<P>Private Sub tcpServer_ConnectionRequest _</P>
<P>(ByVal requestID As Long)</P>
<P>  ' Check if the control's State is closed. If not, </P>
<P>  ' close the connection before accepting the new</P>
<P>  ' connection.</P>
<P>  If tcpServer.State &lt;&gt; sckClosed Then _</P>
<P>  tcpServer.Close</P>
<P>  ' Accept the request with the requestID </P>
<P>  ' parameter.</P>
<P>  tcpServer.Accept requestID</P>
<P>End Sub</P>
<P>Private Sub txtSendData_Change()</P>
<P>  ' The TextBox control named txtSendData </P>
<P>  ' contains the data to be sent. Whenever the user </P>
<P>  ' types into the textbox, the string is sent </P>
<P>  ' using the SendData method.</P>
<P>  tcpServer.SendData txtSendData.Text</P>
<P>End Sub</P>
<P>Private Sub tcpServer_DataArrival _</P>
<P>(ByVal bytesTotal As Long)</P>
<P>  ' Declare a variable for the incoming data. </P>
<P>  ' Invoke the GetData method and set the Text</P>
<P>  ' property of a TextBox named txtOutput to </P>
<P>  ' the data.</P>
<P>  Dim strData As String</P>
<P>  tcpServer.GetData strData</P>
<P>  txtOutput.Text = strData</P>
<P>End Sub</P></DIR>
</DIR>
</FONT><P>The procedures above create a simple server application. However, to complete the scenario, you must also create a client application.</P>
<B><P>To create a TCP client</B> </P><DIR>
<DIR>
<P>Add a new form to the project, and name it frmClient.<BR>
</P>
<P>Change the caption of the form to TCP Client.<BR>
</P>
<P>Add a Winsock control to the form and name it tcpClient.<BR>
</P>
<P>Add two TextBox controls to frmClient. Name the first txtSend, and the second txtOutput.<BR>
</P>
<P>Draw a CommandButton control on the form and name it cmdConnect.<BR>
</P>
<P>Change the caption of the CommandButton control to Connect.<BR>
</P>
<P>Add the code below to the form. </P></DIR>
</DIR>
<B><P>Important</B> Be sure to change the value of the RemoteHost property to the friendly name of your computer.</P>
<FONT FACE="Courier New" SIZE=2><P>Private Sub Form_Load()</P>
<P>  ' The name of the Winsock control is tcpClient.</P>
<P>  ' Note: to specify a remote host, you can use </P>
<P>  ' either the IP address (ex: "121.111.1.1") or</P>
<P>  ' the computer's "friendly" name, as shown here.</P>
<P>  tcpClient.RemoteHost = "RemoteComputerName"</P>
<P>  tcpClient.RemotePort = 1001</P>
<P>End Sub</P>
<P>Private Sub cmdConnect_Click()</P>
<P>  ' Invoke the Connect method to initiate a </P>
<P>  ' connection.</P>
<P>  tcpClient.Connect</P>
<P>End Sub</P>
<P>Private Sub txtSend_Change()</P>
<P>  tcpClient.SendData txtSend.Text</P>
<P>End Sub</P>
<P>Private Sub tcpClient_DataArrival _</P>
<P>(ByVal bytesTotal As Long)</P>
<P>  Dim strData As String</P>
<P>  tcpClient.GetData strData</P>
<P>  txtOutput.Text = strData</P>
<P>End Sub</P>
</FONT><P>The code above creates a simple client-server application. To try the two together, run the project, and click Connect. Then type text into the txtSendData TextBox on either form, and the same text will appear in the txtOutput TextBox on the other form.</P>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>Accepting More than One Connection Request</P>
</B></FONT><P>The basic server outlined above accepts only one connection request. However, it is possible to accept several connection requests using the same control by creating a control array. In that case, you do not need to close the connection, but simply create a new instance of the control (by setting its Index property), and invoking the Accept method on the new instance.</P>
<P>The code below assumes there is a Winsock control on a form named sckServer, and that its Index property has been set to 0; thus the control is part of a control array. In the Declarations section, a module-level variable intMax is declared. In the form's Load event, intMax is set to 0, and the LocalPort property for the first control in the array is set to 1001. Then the Listen method is invoked on the control, making it the "listening control. As each connection request arrives, the code tests to see if the Index is 0 (the value of the "listening" control). If so, the listening control increments intMax, and uses that number to create a new control instance. The new control instance is then used to accept the connection request.</P>
<FONT FACE="Courier New" SIZE=2><P>Private intMax As Long</P>
<P>Private Sub Form_Load()</P>
<P>  intMax = 0</P>
<P>  sckServer(0).LocalPort = 1001</P>
<P>  sckServer(0).Listen</P>
<P>End Sub</P>
<P>Private Sub sckServer_ConnectionRequest _</P>
<P>(Index As Integer, ByVal requestID As Long)</P>
<P>  If Index = 0 Then</P>
<P>   intMax = intMax + 1</P>
<P>   Load sckServer(intMax)</P>
<P>   sckServer(intMax).LocalPort = 0</P>
<P>   sckServer(intMax).Accept requestID</P>
<P>   Load txtData(intMax)</P>
<P>  End If</P>
<P>End Sub</P>
</FONT><B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>UDP Basics</P>
</B></FONT><P>Creating a UDP application is even simpler than creating a TCP application because the UDP protocol doesn't require an explicit connection. In the TCP application above, one Winsock control must explicitly be set to "listen," while the other must initiate a connection with the Connect method.</P>
<P>In contrast, the UDP protocol doesn't require an explicit connection. To send data between two controls, three steps must be completed (on both sides of the connection): </P><DIR>
<DIR>
<P>Set the RemoteHost property to the name of the other computer.<BR>
</P>
<P>Set the RemotePort property to the LocalPort property of the second control.<BR>
</P>
<P>Invoke the Bind method specifying the LocalPort to be used. (This method is discussed in greater detail below.) </P></DIR>
</DIR>
<P>Because both computers can be considered "equal" in the relationship, it could be called a peer-to-peer application. To demonstrate this, the code below creates a "chat" application that allows two people to "talk" in real time to each other:</P>
<B><P>To create a UDP Peer</B> </P><DIR>
<DIR>
<P>Create a new Standard EXE project.<BR>
</P>
<P>Change the name of the default form to frmPeerA.<BR>
</P>
<P>Change the caption of the form to "Peer A."<BR>
</P>
<P>Draw a Winsock control on the form and name it udpPeerA.<BR>
</P>
<P>On the <B>Properties</B> page, click <B>Protocol</B> and change the protocol to UDPProtocol.<BR>
</P>
<P>Add two TextBox controls to the form. Name the first txtSend, and the second txtOutput.<BR>
</P>
<P>Add the code below to the form. </P>
<FONT FACE="Courier New" SIZE=2><P>Private Sub Form_Load()</P>
<P>  ' The control's name is udpPeerA</P>
<P>  With udpPeerA</P>
<P>    ' IMPORTANT: be sure to change the RemoteHost </P>
<P>    ' value to the name of your computer.</P>
<P>    .RemoteHost= "PeerB" </P>
<P>    .RemotePort = 1001  ' Port to connect to.</P>
<P>    .Bind 1002        ' Bind to the local port.</P>
<P>  End With</P>
<P>  frmPeerB.Show         ' Show the second form.</P>
<P>End Sub</P>
<P>Private Sub txtSend_Change()</P>
<P>  ' Send text as soon as it's typed.</P>
<P>  udpPeerA.SendData txtSend.Text</P>
<P>End Sub</P>
<P>Private Sub udpPeerA_DataArrival _</P>
<P>(ByVal bytesTotal As Long)</P>
<P>  Dim strData As String</P>
<P>  udpPeerA.GetData strData</P>
<P>  txtOutput.Text = strData</P>
<P>End Sub</P></DIR>
</DIR>
</FONT><B><P>To create a second UDP Peer</B> </P><DIR>
<DIR>
<P>Add a standard form to the project.<BR>
</P>
<P>Change the name of the form to frmPeerB.<BR>
</P>
<P>Change the caption of the form to "Peer B."<BR>
</P>
<P>Draw a Winsock control on the form and name it udpPeerB.<BR>
</P>
<P>On the <B>Properties</B> page, click <B>Protocol</B> and change the protocol to UDPProtocol.<BR>
</P>
<P>Add two TextBox controls to the form. Name the TextBox txtSend, and the second txtOutput.<BR>
</P>
<P>Add the code below to the form. </P>
<FONT FACE="Courier New" SIZE=2><P>Private Sub Form_Load()</P>
<P>  ' The control's name is udpPeerB.</P>
<P>  With udpPeerB</P>
<P>    ' IMPORTANT: be sure to change the RemoteHost </P>
<P>    ' value to the name of your computer.</P>
<P>    .RemoteHost= "PeerA"</P>
<P>    .RemotePort = 1002  ' Port to connect to.</P>
<P>    .Bind 1001        ' Bind to the local port.</P>
<P>  End With</P>
<P>End Sub</P>
<P>Private Sub txtSend_Change()</P>
<P>  ' Send text as soon as it's typed.</P>
<P>  udpPeerB.SendData txtSend.Text</P>
<P>End Sub</P>
<P>Private Sub udpPeerB_DataArrival _</P>
<P>(ByVal bytesTotal As Long)</P>
<P>  Dim strData As String</P>
<P>  udpPeerB.GetData strData</P>
<P>  txtOutput.Text = strData</P>
<P>End Sub</P></DIR>
</DIR>
</FONT><P>To try the example, press F5 to run the project, and type into the txtSend TextBox on either form. The text you type will appear in the txtOutput TextBox on the other form.</P>
<B><FONT SIZE=5><P>|----------------------------------------------------------------------|</P>
<P>About the Bind Method</P>
</B></FONT><P>As shown in the code above, you must invoke the Bind method when creating a UDP application. The Bind method "reserves" a local port for use by the control. For example, when you bind the control to port number 1001, no other application can use that port to "listen" on. This may come in useful if you wish to prevent another application from using that port.</P>
<P>The Bind method also features an optional second argument. If there is more than one network adapter present on the machine, the <I>LocalIP </I>argument allows you to specify which adapter to use. If you omit the argument, the control uses the first network adapter listed in the Network control panel dialog box of the computer's Control Panel Settings.</P>
<P>When using the UDP protocol, you can freely switch the RemoteHost and RemotePort properties while remaining bound to the same LocalPort. However, with the TCP protocol, you must close the connection before changing the RemoteHost and RemotePort properties.</P></BODY>
</HTML>

