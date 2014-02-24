# TODO: make functions man-able
# TODO: Error checking
# TODO: Make service accounts truly gen-able by error checking
# TODO: Reset password functionality

$script:secrets = @{
                "dev"  = @{"key"="d3v_p@ssphr@s3";"iv"="d3v_v3ct0r";"server"="sdvdc1";"suffix"="dev.oeconnection.com";"baseDN"="DC=dev,DC=oeconnection,DC=com"};
                "qa"   = @{"key"="q@_p@ssphr@s3";"iv"="q@_v3ct0r";"server"="sqadc1";"suffix"="qa.oeconnection.com";"baseDN"="DC=qa,DC=oeconnection,DC=com"};
                "prod" = @{"key"="pr0d_p@ssphr@s3";"iv"="pr0d_v3ct0r";"server"="sproddc1";"suffix"="prod.oeconnection.com";"baseDN"="DC=prod,DC=oeconnection,DC=com"};
                "oec"  = @{"key"="03c_p@ssphr@s3";"iv"="03c_v3ct0r";"server"="sisdc1";"suffix"="oec.oeconnection.com";"baseDN"="DC=oec,DC=oeconnection,DC=com"};
                "svvs" = @{"key"="svvs_p@ssphr@s3";"iv"="svvs_v3ct0r";"server"="s265575ch3vw21";"suffix"="s265575-ad01.corp";"baseDN"="DC=s265575-ad01,DC=corp"}
}

function reverse([string]$string) {
    for ($i = $string.length - 1; $i -ge 0; $i--) {$ns = $ns + ($string.substring($i,1))}
    $ns
}

function encrypt([string]$string,[string]$key,[string]$iv) {
                $rij = New-Object System.Security.Cryptography.RijndaelManaged
                $keyAsBytes = [System.Text.Encoding]::ASCII.GetBytes($key)
                $ivAsBytes = [System.Text.Encoding]::ASCII.GetBytes($iv)
                $fixedKeyAsBytes = new-object byte[] 16
                $fixedIvAsBytes = new-object byte[] 16
                [Array]::Copy($keyAsBytes,$fixedKeyAsBytes,$keyAsBytes.Length)
                [Array]::Copy($ivAsBytes,$fixedIvAsBytes,$ivAsBytes.Length)
                $encryptor = $rij.CreateEncryptor($fixedKeyAsBytes,$fixedIvAsBytes)

                $ms = new-Object IO.MemoryStream
                $cs = new-Object Security.Cryptography.CryptoStream $ms,$encryptor,"Write" # Target data stream, transformation, and mode
                $sw = new-Object IO.StreamWriter $cs
                $sw.Write($string)          # Write the string through the crypto stream into the memory stream
                $sw.Close()
                $cs.Close()
                $ms.Close()
                $rij.Clear()
                [byte[]]$result = $ms.ToArray()                # Byte array from the encrypted memory stream
                [Convert]::ToBase64String($result)         # Convert to base64 for transport
}

function CreateAccount([string]$accountType,[string]$uname,[string]$upwd,[string]$description,[string]$env,[string]$ftpiis, [string]$iispath) {
                switch ($accountType) {
                                "ftp" { $ou = 'OU=FTP Accounts,OU=Accounts'; $group = 'CN=FTPAccounts,OU=Groups'; $ftpiis; $iispath}
                                "service" { $ou = 'OU=Service Accounts,OU=Accounts'; $group = 'CN=ServiceAccounts,OU=Groups'}
                }
                $adsi = [ADSI]"LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$ou,$($secrets.$env.baseDN)"
                $cn = "CN="+ $uname
                $user = $adsi.Create("User", $cn)
                $user.setInfo()
                $user = [ADSI]"LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$cn,$ou,$($secrets.$env.baseDN)"
                $user.Put("displayName", $uname)
                $user.Put("sAMAccountName", $uname)
                $user.Put("userPrincipalName", "$uname`@$($secrets.$env.suffix)")
                $user.Put("description", "$description")
                $user.SetInfo()
                $user.SetPassword($upwd)
                $user.setInfo()
				
				[int]$UF_ACCOUNTDISABLE = 0x0002;
				[int]$UF_PASSWD_NOTREQD = 0x0020;
				[int]$UF_PASSWD_CANT_CHANGE = 0x0040;
				[int]$UF_NORMAL_ACCOUNT = 0x0200;
				[int]$UF_DONT_EXPIRE_PASSWD = 0x10000;
				[int]$UF_SMARTCARD_REQUIRED = 0x40000;
				[int]$UF_PASSWORD_EXPIRED = 0x800000;
				
				#enable account and don't expire password
				[int]$userControlFlags = $UF_NORMAL_ACCOUNT + $UF_DONT_EXPIRE_PASSWD
				$user.Properties["userAccountControl"].Value = $userControlFlags
				$user.CommitChanges()
			    
				#user can't change password
				$self = [System.Security.Principal.SecurityIdentifier]'S-1-5-10' 
				$SelfDeny = new-object System.DirectoryServices.ActiveDirectoryAccessRule($self,'ExtendedRight','Deny',[Guid]'ab721a53-1e2f-11d0-9819-00aa0040529b') 
				$user.psbase.get_ObjectSecurity().AddAccessRule($SelfDeny) 
				$user.psbase.CommitChanges()
				
				if ($accountType -eq "service" -and $env -eq "oec") 
				{
					#Don't add user to group
				}
				else 
				{
					$Group = [ADSI]"LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$group,$($secrets.$env.baseDN)"
					$userToAdd = "LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$cn,$ou,$($secrets.$env.baseDN)"
					$Group.Add($usertoadd)
				}
				
				if ($ftpIIS -and $iispath)
				{
					ManageWebs $env addftpuser uns_ftp_nodes $uname $iispath
				}			
}

function CreateServiceAccount() {
				param 
				(
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$uname
					,
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$description
					,
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$env
					,
					[Parameter(Mandatory=$false,ValueFromPipeline=$true)] 
					[boolean]$gen
				)
				
				$env = $env.ToLower()
				[string]$i=1
                if ($gen) { 
                                $hour = ([datetime]::Now.Hour).ToString("00")
                                [char]$idx='a'
                                $uname = "s$uid$hour$idx".ToLower()
                } 
                else {}
                $unEncPwd = "$i$(reverse $uname)_$env"
                $description = $description + " : [i$i]"
                $encPwd = (encrypt $unEncPwd $secrets.$env.key $secrets.$env.iv)
                if ($env -ne "all")
				{
					CreateAccount "service" $uname $encPwd $description $env
				}
				else 
				{
					foreach ($env in $script:secrets)
					{
						CreateAccount "service" $uname $encPwd $description $env
					}
				}
                "Environment: $env User: $uname Password: $encPwd"
}

function CreateFtpAccount() {
                Param 
				(
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$uname
					, 
					[Parameter(Mandatory=$false,ValueFromPipeline=$true)] 
					[string]$upwd = $null
					,
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$description
					,
					[Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
					[string]$env
					,
					[Parameter(Mandatory=$false,ValueFromPipeline=$true)] 
					[boolean]$iisftp
					,
					[Parameter(Mandatory=$false,ValueFromPipeline=$true)] 
					[string]$iispath
				)
				$env = $env.ToLower()
				[string]$i=1
				if (-not $upwd) {
                                $unEncPwd = "$i$(reverse $uname)_$env"
                                $description = $description + " : [i$i]"
                                $encPwd = (encrypt $unEncPwd $secrets.$env.key $secrets.$env.iv)
                                $upwd = $encPwd
                }
				
                CreateAccount "ftp" $uname $upwd $description $env $iisftp $iispath
                "Environment: $env User: $uname Password: $upwd "
}

function RetrievePwd([string]$uname,[string]$env,[string]$i=1) {
  $env = $env.ToLower()
  $unEncPwd = "$i$(reverse $uname)_$env"
  $encPwd = (encrypt $unEncPwd $secrets.$env.key $secrets.$env.iv)
  "User: $uname Password: $encPwd"
}

function CreateGroup([string]$groupname,[string]$env,[string]$desc,[string]$scope) {
                $env = $env.ToLower()
				switch ($scope) {
                                "domain local" { $groupType = "-2147483644" }
                                default { $groupType = "-2147483644" }
                }
                $ou = 'OU=Groups'
                $adsi = [ADSI]"LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$ou,$($secrets.$env.baseDN)"
                $cn = "CN="+ $groupname
                $group = $adsi.Create("Group", $cn)
                $group.Put("groupType", $groupType)
                $group.setInfo()
                $group = [ADSI]"LDAP://$($secrets.$env.server).$($secrets.$env.suffix):389/$cn,$ou,$($secrets.$env.baseDN)"
                $group.Put("sAMAccountName", $groupname)
                if ($desc) { $group.Put("description", "$description") }
                $group.setInfo()
}

function ModifyDNSRecord([string]$env,[string]$nodeName,[string]$rrType,[string]$rrData,[string]$task,[string]$server,[string]$zone) {
                $env = $env.ToLower()
				if ($server -eq $null) { $server = $($secrets.$env.server + '.' + $secrets.$env.suffix) }
                if ($zone -eq $null) { $zone =  $($secrets.$env.suffix) }
                if ($task -eq "add") { & dnscmd.exe $server /RecordAdd $zone $nodeName $rrType $rrData }
                if ($task -eq "del") { & dnscmd.exe $server /RecordDelete $zone $nodeName $rrType $rrData /f }
}

function SwapDNS([string]$loc) {
                if ($loc -eq "savvis") {
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 10.47.121.50
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 10.47.121.51
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 172.17.27.1
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 172.17.27.2
                }
                if ($loc -eq "oec") {
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 172.17.27.1
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 172.17.27.2
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 10.47.121.50
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 10.47.121.51
                }
  if ($loc -eq "qa") {
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 10.1.21.21
                                & netsh interface ipv4 add dnsserver name="Local Area Connection" 10.1.21.22
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 172.17.27.1
                                & netsh interface ipv4 delete dnsserver name="Local Area Connection" 172.17.27.2
                }
                & ipconfig /flushdns
}

Export-ModuleMember -Function encrypt,reverse,CreateServiceAccount,RetrievePwd,CreateGroup,ModifyDNSRecord,CreateServiceAccount,CreateFtpAccount,CreateAccount,SwapDNS
Export-ModuleMember -Variable $script:secrets

#CreateServiceAccount -uname ryantest -description "Ryan test account" -env all
#CreateFTPAccount -uname ryantest -description "ryantest ftp account" -env qa -iisftp $true -iispath E:\Users\ryantest  
