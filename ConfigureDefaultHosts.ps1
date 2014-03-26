# Copyright 2012 Scott Banwart

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

# http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# BizTalk host configuration script
# Version 1.0

# BizTalk WMI enumeration constants
$HOST_IN_PROCESS = 1
$HOST_ISOLATED = 2

# Default BizTalk in-process host
$DEFAULT_HOST = "BizTalkServerApplication"

# Configuration options

# Server name
$SERVER_NAME = ""

# Default BizTalk application users group
$NT_GROUP_NAME = ''

# Default BizTalk service credentials
$IN_PROCESS_USERNAME = ''
$IN_PROCESS_PASSWORD = ""

$ISOLATED_HOST_USERNAME = $IN_PROCESS_USERNAME
$ISOLATED_HOST_PASSWORD = $IN_PROCESS_PASSWORD

function Write-HostSuccess($message)
{
    Write-Host -ForegroundColor Green $message
}

function Write-HostError($message)
{
    Write-Host -ForegroundColor Red $message
}

$hostsToAdd = @(
    @{Name='BizTalkSendHost';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$False;HostTracking=$False}
   ,@{Name='BizTalkReceiveHost';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$False;HostTracking=$False}
   ,@{Name='BizTalkSend32Host';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$True;HostTracking=$False}
   ,@{Name='BizTalkReceive32Host';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$True;HostTracking=$False}
   ,@{Name='BizTalkOrchestrationHost';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$False;HostTracking=$False}
   ,@{Name='BizTalkTrackingHost';NTGroupName="$NT_GROUP_NAME";AuthTrusted=$False;HostType=$HOST_IN_PROCESS;IsHost32BitOnly=$False;HostTracking=$True}

)

$hostInstancesToAdd = @(
     @{HostName='BizTalkSendHost';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
    ,@{HostName='BizTalkReceiveHost';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
    ,@{HostName='BizTalkSend32Host';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
    ,@{HostName='BizTalkReceive32Host';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
    ,@{HostName='BizTalkOrchestrationHost';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
    ,@{HostName='BizTalkTrackingHost';Username=$IN_PROCESS_USERNAME;Password=$IN_PROCESS_PASSWORD;ServerName=$SERVER_NAME}
)

$receiveHandlers = @(
     @{AdapterName='FILE';HostName='BizTalkReceiveHost'}
    ,@{AdapterName='FTP';HostName='BizTalkReceive32Host'}
    ,@{AdapterName='MQSeries';HostName='BizTalkReceiveHost'}
    ,@{AdapterName='MSMQ';HostName='BizTalkReceiveHost'}
    ,@{AdapterName='POP3';HostName='BizTalkReceive32Host'}
    ,@{AdapterName='SB-Messaging';HostName='BizTalkReceiveHost'}
    ,@{AdapterName='SFTP';HostName='BizTalkReceive32Host'}
    ,@{AdapterName='SQL';HostName='BizTalkReceive32Host'}
    ,@{AdapterName='Windows SharePoint Services';HostName='BizTalkReceive32Host'}
)

$sendHandlers = @(
     @{AdapterName='FILE';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='FTP';HostName='BizTalkSend32Host';IsDefault=$True}
    ,@{AdapterName='HTTP';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='HTTP';HostName='BizTalkSend32Host';IsDefault=$False}
    ,@{AdapterName='MQSeries';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='MSMQ';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='SB-Messaging';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='SFTP';HostName='BizTalkSend32Host';IsDefault=$True}
    ,@{AdapterName='SMTP';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='SOAP';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='SQL';HostName='BizTalkSend32Host';IsDefault=$True}
    ,@{AdapterName='WCF-BasicHttp';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-BasicHttpRelay';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-BasicCustom';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-BasicCustomRelay';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-NetMsmq';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-NetNamedPipe';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-NetTcp';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-NetTcpRelay';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-SQL';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-WebHttp';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='WCF-WSHttp';HostName='BizTalkSendHost';IsDefault=$True}
    ,@{AdapterName='Windows SharePoint Services';HostName='BizTalkSend32Host';IsDefault=$True}
)

foreach ($hostParams in $hostsToAdd)
{
    try
    {
        $newHost = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostSetting").CreateInstance()
        $newHost["Name"] = $hostParams["Name"]
        $newHost["NTGroupName"] = $hostParams["NTGroupName"]
        $newHost["AuthTrusted"] = $hostParams["AuthTrusted"]
        $newHost["HostType"] = $hostParams["HostType"]
        $newHost["IsHost32BitOnly"] = $hostParams["IsHost32BitOnly"]
        $newHost["HostTracking"] = $hostParams["HostTracking"]
        $newHost.Put()
        
        $name = $hostParams["Name"]
        Write-HostSuccess "Successfully created host: $name"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        $name = $hostParams["Name"]
        Write-HostError "An error occurred while adding host: $name, $_.Exeception.ToString()"
    }
}

foreach ($instanceParams in $hostInstancesToAdd)
{
    try
    {
        $serverHost = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_ServerHost").CreateInstance()
        $serverHost["HostName"] = $instanceParams["HostName"]
        $serverHost["ServerName"] = $instanceParams["ServerName"]
        $serverHost.Map()

        $hostInstance = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostInstance").CreateInstance()
        $name = "Microsoft BizTalk Server " + $instanceParams["HostName"] + " " + $instanceParams["ServerName"]
        $hostInstance["Name"] = $name
        $hostInstance.Install($instanceParams["Username"], $instanceParams["Password"], $True)

        Write-HostSuccess "Successfully created host instance: $name"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        Write-HostError "An error occurred while adding host instance: $name, $_.Exeception.ToString()"
    }
}

foreach ($handler in $receiveHandlers)
{
    try
    {
        $hostName = $handler["HostName"]
        $adapterName = $handler["AdapterName"]

        $receiveHandler = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_ReceiveHandler").CreateInstance()
        $receiveHandler["AdapterName"] = $handler["AdapterName"]
        $receiveHandler["HostName"] = $handler["HostName"]
        $receiveHandler.Put()

        
        $receiveHandler = Get-WmiObject 'MSBTS_ReceiveHandler' -Namespace 'root\MicrosoftBizTalkServer' -Filter "HostName='$DEFAULT_HOST' AND AdapterName='$adapterName'"
        $receiveHandler.Delete()

        Write-HostSuccess "Successfully created receive handler on host: $hostName for adapter: $adapterName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        $hostName = $handler["HostName"]
        $adapterName = $handler["AdapterName"]
        Write-HostError "An error occurred while adding receive hander on host: $hostName for adapter: $adapterName, $_.Exeception.ToString()"
    }
}

foreach ($handler in $sendHandlers)
{
    try
    {
        $hostName = $handler["HostName"]
        $adapterName = $handler["AdapterName"]

        $sendHandler = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_SendHandler2").CreateInstance()
        $sendHandler["AdapterName"] = $handler["AdapterName"]
        $sendHandler["HostName"] = $handler["HostName"]
        $sendHandler["IsDefault"] = $handler["IsDefault"]
        $sendHandler.Put()

        $sendHandler = Get-WmiObject 'MSBTS_SendHandler2' -Namespace 'root\MicrosoftBizTalkServer' -Filter "HostName='$DEFAULT_HOST' AND AdapterName='$adapterName'"
        $sendHandler.Delete()

        Write-HostSuccess "Successfully created send handler on host: $hostName for adapter: $adapterName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        $hostName = $handler["HostName"]
        $adapterName = $handler["AdapterName"]
        $errMsg = $_.Exception.ToString()
        Write-HostError "An error occurred while adding send handler on host: $hostName for adapter: $adapterName, $errMsg"
    }
}

function Write-HostSuccess($message)
{
    Write-Host -ForegroundColor Green $message
}

function Write-HostError($message)
{
    Write-Host -ForegroundColor Red $message
}
