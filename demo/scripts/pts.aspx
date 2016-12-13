<%@ Page Language="VB" AutoEventWireup="false" CodeFile="pts.aspx.vb" Inherits="Default3" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Patient Tracking System</title>
<base target="rtop">
</head>

<body>


<form id="form1" runat="server" target="main.html" title="Home">


<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="Batang">Welcome to the the Vancouver Island Health Authority patient 
    tracking system!</font></p>
	<p>
        <asp:AccessDataSource ID="AccessDataSource1" runat="server" 
            DataFile="/data/group7DB.mdb" 
            SelectCommand="SELECT Admission.[EncounterNum], Patient.PHN, Patient.LastName, Admission.[AdmissionDate], Location.[UnitID], Location.[RoomNum] FROM ((Admission INNER JOIN Patient ON Admission.PHN = Patient.PHN) INNER JOIN Location ON Admission.[EncounterNum] = Location.[EncounterNum])
            WHERE (((Admission.[EncounterNum])= @Title))">
             <SelectParameters>
            <asp:ControlParameter Name="Title" 
      ControlID="TextBox1"/>
  </SelectParameters>
        </asp:AccessDataSource>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</p>
	<p><font face="Batang">Please enter encounter number&nbsp; </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="Submit"  style="height: 26px"/>
</p>
<p>&nbsp;</p>
<p>
    <asp:GridView ID="GridView1" runat="server" DataSourceID="AccessDataSource1">
    </asp:GridView>
</p>
<p align="center">
    &nbsp;</p>
<p align="center">
    &nbsp;</p>
<p align="center">
    &nbsp;</p>
<p align="center">
<a href="http://www.viha.ca">
<img border="0" src="/images/viha-img.jpg" width="286" height="109"></a></p>

</form>
&gt;<a target="_top" href="/index.html">Home</a></font></p>

</body>

</html>
