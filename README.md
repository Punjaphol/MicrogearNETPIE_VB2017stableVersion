# MicrogearNETPIE_VB2017stableVersion
'Create by Mr. Punjaphol Kengna / Software developper /Nidec-Shibaura 'Copyright ©  2017 'THis software distributed under GNU license. 'Author-email:Punjaphol-K@nidec-shibaura.co.th

MicroGearVB2017byPunjaphol.dll

เป็น Dll สำหรับ ดึง/ส่ง ข้อมูล Netpie ตัวใหม่ สร้างจาก VB ล้วน
สามารถ Import มาใช้งานได้เลยเป็นแบบ Event drivent. (เรียก Control อื่นได้โดยตรง)

แก้ปัญหา  Microgear.DLL ตัวเก่าversion c#
-1.กิน CPU 40% ตลอด
-2.รันแอพค้างคืนมักจะหลุด แล้วไม่ต่อใหม่ ทำให้ขาดความต่อเนื่องในการเชื่อมต่อ..ต้องคอยปิดเปิดโปรใหม่


**ปล.Import***
M2Mqtt.NetCf39.dll
MicroGearVB2017byPunjaphol.dll
RestSharp.dll
ก่อนใช้งาน







Public Class Form1

    Dim APP_ID As String = "APPTESTNETPIE" ' "APP_ID" ' "APP_ID"
    Dim APP_KEY As String = "E562i71Tqy79eAF" '"APP_KEY"
    Dim APP_SECRET As String = "ChOLxvZVDHMlCM7lXD9IrncaC" '"APP_SECREAT"
    Dim NetpieAlias As String = "PCclient"
    Dim topic As String = "/sensordata" ' "/MyDataTopic"

    Dim WithEvents Netpie2 As New MicroGearVB2017byPunjaphol.MicroGearVB2017byPunjaphol

    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Netpie2.Disconnect()
        Netpie2.ResetToken()
        End
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Netpie2.Connect(APP_ID, APP_KEY, APP_SECRET)
        TextBox1.Text = "Connecting.."
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub Netpie2_onAbsent(token As String) Handles Netpie2.onAbsent
        Console.WriteLine(token)
    End Sub

    Private Sub Netpie2_onConnect() Handles Netpie2.onConnect
        Button1.BackColor = Color.Green
        TextBox1.Text = "PIE CONNECTED"
        Console.WriteLine("Connected netpie")
        Button1.Enabled = True
        Button2.Enabled = True
        Netpie2.SetAlias(NetpieAlias)
        Netpie2.Subscribe(topic) 'Register to recieve the message from topic group

    End Sub

    Private Sub Netpie2_onDisconnect() Handles Netpie2.onDisconnect
        Console.WriteLine("DisConnected netpie")
    End Sub

    Private Sub Netpie2_onError(Message As String) Handles Netpie2.onError
        Try
            Console.WriteLine(Message)
            Netpie2.ResetToken() 'reset token alway error
            Netpie2.Connect(APP_ID, APP_KEY, APP_SECRET)
            TextBox1.Text = "Connecting.."
            Button1.Enabled = False
            Button2.Enabled = False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Private Sub Netpie2_onInfo(txtInfo As String) Handles Netpie2.onInfo
        Console.WriteLine(txtInfo)
    End Sub


    Private Sub Netpie2_onMessage(topic As String, message As String) Handles Netpie2.onMessage

        Console.WriteLine(topic & "/" & message)
        TextBox1.Text = message

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Netpie2.writeFeed("Test32Field", "{Field1:5,Field2:10,Field3:15,Field4:20}")

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Netpie2.Publish(topic, TextBox2.Text)
    End Sub

    Private Sub Netpie2_onPresent(token As String) Handles Netpie2.onPresent
        Console.WriteLine(token)
    End Sub
End Class


