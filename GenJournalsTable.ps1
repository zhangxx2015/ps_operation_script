$comments = @'
zhangxx 2017-07-31
'@

$CurDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$CurPs = $MyInvocation.MyCommand.Definition.Replace($CurDir.ToString(),"")
$config = @{}
$path = $CurDir,$CurPs,".conf" -join ""
$payload = Get-Content -encoding utf8 -Path $path | Where-Object { $_ -like '*:=*' } | ForEach-Object {
	$infos = $_ -split ':='
	$key = $infos[0].Trim()
    $value = $infos[1].Trim()
    $config.$key = $value
}
$Today = Get-Date -Format 'yyyy-MM-dd'
# 输出文件 默认为:当前目录\当文件名_当前日期.csv
$ExportCSV = $CurDir,$CurPs,'_',$Today,".csv" -join ""

$MailBody = $config.MailBody.ToString().Replace("{today}",$Today)
$Subject  = $MailBody
$SQlQuery = $config.SQlQuery
$ConfigIndex  = $config.ConfigIndex
$ConnStr  = ($config.ConnStr0,$config.ConnStr1)[$ConfigIndex].ToString()


# Smtp 设置
$SmtpServer 	= $config.SmtpServer
$EnableSsl 		= $config.EnableSsl
$SmtpServerPort = $config.SmtpServerPort
$SmtpUser 		= $config.SmtpUser
$SmtpPass 		= $config.SmtpPass
$SmtpTimeout 	= $config.SmtpTimeout

# 邮件设置
$FromName     	= ($config.FromName0,$config.FromName1)[$ConfigIndex].ToString()
$To 			= ($config.To0,$config.To1)[$ConfigIndex].ToString()
$Cc				= ($config.Cc0,$config.Cc1)[$ConfigIndex].ToString()

'global config-----------------------------------------'
$SmtpServer
$EnableSsl
$SmtpServerPort
$SmtpUser
$SmtpPass
$SmtpTimeout

$ExportCSV
$MailBody
$SQlQuery
'custom config-----------------------------------------'
$ConfigIndex
$ConnStr

$FromName
$To
$Cc


#[Console]::WriteLine("done")
#[Console]::ReadLine()
#return


	Try {
		[system.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\MySQL\MySQL Connector Net 5.0.9\Binaries\.NET 2.0\MySql.Data.dll") | Out-Null
		[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
		$Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
		$Connection.ConnectionString = $ConnStr
		$Connection.Open()
		# $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $Connection)
		$Command = New-Object MySql.Data.MySqlClient.MySqlCommand
		$Command.Connection = $Connection
		$Command.CommandText = $SQlQuery
		$Command.CommandTimeout = 300000
		
		$DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
		$DataSet = New-Object System.Data.DataSet
		$RecordCount = $dataAdapter.Fill($dataSet, "data")
		$table = $DataSet.Tables["data"]
		$table | Export-CSV $ExportCSV -Encoding utf8
Write-Output "done."
Start-Sleep -Seconds 5
return


		$From 			= $SmtpUser
		$HasAttachment 	= $true
		$AttachmentPath = $ExportCSV
		$IsBodyHtml		= $false
		$Body			= $MailBody
		# /////////////////////////////////////////////////////////////////////////////////////////
		[System.Reflection.Assembly]::LoadWithPartialName("System.Web") > $null
		$mail = New-Object System.Web.Mail.MailMessage
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", 				$SmtpServer)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", 			$SmtpServerPort)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", 				$EnableSsl)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", 			$SmtpUser)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", 			$SmtpPass)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout", 	$SmtpTimeout / 1000)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", 				2)
		$mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", 		1)
		$mail.From		= $From
		$mail.To		= $To
		$mail.Cc		= $Cc
		$mail.Subject	= $Subject
		$mail.Body		= $Body
		if ($HasAttachment){
			$AttachmentPath = (get-item $AttachmentPath).FullName
			$attachment 	= New-Object System.Web.Mail.MailAttachment $AttachmentPath
			$mail.Attachments.Add($attachment) > $null
		}
		Write-Output "Sending email to $to..."
		try{
			[System.Web.Mail.SmtpMail]::Send($mail)
			Write-Output "Message sent."
		}catch{
			Write-Error $_
			Write-Output "Message send failed."
		}


	} Catch {
		Write-Host "ERROR : Unable to run query : $Sql `n$Error[0]"
	} Finally {
		$Connection.Close()
	}


Start-Sleep -Seconds 5
#[Console]::WriteLine("done")
#[Console]::ReadLine()
