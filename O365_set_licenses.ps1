#Set licenses Office 365 users.
#Date
$date = date -Format "dd.MM.yyyy HH:mm:ss"
$date_logs = date -Format "dd.MM.yyyy"
$date

function to_log ([string]$data_log)
 {
$date = date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content C:\Logs\set-licenses_$date_logs.log "$date, $data_log"

}

function to_log_sql ([string]$data_log_sql)
 {
$date = date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content C:\Logs\SQL_set-licenses_$date_logs.log "$date, $data_log_sql"

}

try {
#Import Module SQL Server
Import-Module sqlserver -ErrorAction Stop
to_log "Debug. Модуль SQL Server импортирован успешно"
}
catch {
to_log "Error. Не удалось импортировть модуль SQL Server. $_" -ErrorAction Stop
exit
}
#Connect to O365 
$pass1 = ""
$admin_o365 = ""
$pass = ConvertTo-SecureString -string $pass1 -asplaintext -force
$msolcred = New-Object System.Management.Automation.PSCredential $admin_o365, $pass
try {
connect-msolservice -credential $msolcred -ErrorAction Stop
to_log "Debug. Соединение с О365 - OK"
}
catch {

to_log "Error. Не удалось установить соединение с О365. $_" -ErrorAction Stop
exit
}

#Connect to Exchange Online
try {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $msolcred -Authentication Basic -AllowRedirection -ErrorAction Stop
to_log "Debug. Соединение c Exchange Online - OK"
}
catch {

to_log "Error. Не удалось установить соединение с Exchange Online. $_" -ErrorAction Stop

}

#Cred for SQL
$SQL_Server = ''
$database = ''
$database_user = ''
$database_user_pass = ''


#Set products
try {
$LO = New-MsolLicenseOptions -AccountSkuId "" -ErrorAction Stop #Нет отключенных продуктов 
$LO_NoExchange = New-MsolLicenseOptions -AccountSkuId "" -DisabledPlans "EXCHANGE_S_STANDARD" -ErrorAction Stop #Включено все, кроме Exchange
to_log "Debug. Задание продуктов с отключенным и включенным Exchange - OK"
}
catch {
to_log  "Error. Не удалось задать продукты с отключенным и включенным Exchange. $_" -ErrorAction Stop
exit
}


function sql_query ([string]$sql_query_string) {
try {
Invoke-SQLCmd -ServerInstance $SQL_Server -database $database -username $database_user -password $database_user_pass -Query "$sql_query_string" -ErrorAction "Continue"
    #to_log_sql "Debug. Запрос в БД: $sql_query_string" #закомментил для уменьшения объема логов
    }
catch { 
    
    to_log_sql "Error. Запрос в БД не выполнен. Запрос $sql_query_string. $_"
    
    }
}


function Update_TotalTable_Student_lic ([string]$Change_val) {
    if ($user_in_sql -ne $null) { #если пользователь есть в таблице Total
    if (($user_in_sql.LicenseName -like "") -or ($user_in_sql.LicenseName -like "") -or ($user_in_sql.LicenseName -like $null) ) { 
        sql_query "UPDATE TotalTable SET LicenseName = '', LicenseDate = '$date'  WHERE SAMAccountName like '$samaccountname'"
        sql_query "INSERT INTO LicensesChangeTable (SAMAccountName,LicenseName,Date,Change) VALUES('$samaccountname','','$date','$Change_val')"   
       }
    }
}

function Update_TotalTable_Staff_lic ([string]$Change_val) {
    if ($user_in_sql -ne $null) { #если пользователь есть в таблице Total
    if (($user_in_sql.LicenseName -like "_:STANDARDWOFFPACK_IW_STUDENT") -or ($user_in_sql.LicenseName -like "_:Both") -or ($user_in_sql.LicenseName -like $null)) { 
        sql_query "UPDATE TotalTable SET LicenseName = '_:STANDARDWOFFPACK_IW_FACULTY', LicenseDate = '$date'  WHERE SAMAccountName like '$samaccountname'"
        sql_query "INSERT INTO LicensesChangeTable (SAMAccountName,LicenseName,Date,Change) VALUES('$samaccountname','_:STANDARDWOFFPACK_IW_FACULTY','$date','$Change_val')"   
       }
    }
}

function Update_TotalTable_Both_lic ([string]$Change_val) {
    if ($user_in_sql -ne $null) { #если пользователь есть в таблице Total
    if (($user_in_sql.LicenseName -like "_:STANDARDWOFFPACK_IW_STUDENT") -or ($user_in_sql.LicenseName -like "_:STANDARDWOFFPACK_IW_FACULTY") -or ($user_in_sql.LicenseName -like $null)) { 
        sql_query "UPDATE TotalTable SET LicenseName = '_:Both', LicenseDate = '$date'  WHERE SAMAccountName like '$samaccountname'"
        sql_query "INSERT INTO LicensesChangeTable (SAMAccountName,LicenseName,Date,Change) VALUES('$samaccountname','_:Both','$date','$Change_val')"   
       }
    }
}


function Update_TotalTable_OfficeSync {
     if ($user_in_sql -ne $null) { #если пользователь есть в таблице Total
        if ($user_in_sql.OfficeSync -notlike $date_last_sync) {
        sql_query "UPDATE TotalTable SET OfficeSync = '$date_last_sync'  WHERE SAMAccountName like '$samaccountname'"
       } 
      } 
}

function Update_UseProductTable { 
    #сделать ещё раз запрос по лицензиям
    if ($user_in_sql -ne $null) { #если пользователь есть в таблице Total
        $o365user = Get-MsolUser -UserPrincipalName $UPN #Еще раз делаем запрос по пользователю, чтобы отразить обновленную информацию по нему
        $num = $o365user.Licenses.ServiceStatus.Count - 1
        $UseProductTable_account = sql_query "select * from UserProductTable where SAMAccountName like '$samaccountname'"
    
        #Заполнить, если пользователя ещё не было
        if ($UseProductTable_account -eq $null) { 
    
    
     for ($i=0; $i -le $num; $i++) {
        $ServiceName = $o365user.Licenses.ServiceStatus[$i].ServicePlan.ServiceName
        $ProvisioningStatus = $o365user.Licenses.ServiceStatus[$i].ProvisioningStatus
        sql_query "INSERT INTO UserProductTable (SAMAccountName,ProductName,Date,Change,IsActive) VALUES('$samaccountname','$ServiceName','$date','Данные добавлены','$ProvisioningStatus')"  


            }
            }
    
    #Добавить продукт, если он появился в О365
     $UseProductsTable_diff = Compare-Object -referenceobject $o365user.Licenses.ServiceStatus.ServicePlan.ServiceName -DifferenceObject $UseProductTable_account.ProductName
        foreach ($UseProductTable_diff in $UseProductsTable_diff) {
        for ($i=0; $i -le $num; $i++) {
        $o365user_serviceName =  $o365user.Licenses.ServiceStatus[$i].ServicePlan | where ServiceName -Like $UseProductTable_diff.InputObject

        if ($o365user_serviceName.ServiceName -ne $null) {
        $o365user_provstat =  $o365user.Licenses.ServiceStatus[$i].ProvisioningStatus
        $o365user_serviceName_ = $o365user_serviceName.ServiceName
        sql_query "INSERT INTO UserProductTable (SAMAccountName,ProductName,Date,Change,IsActive) VALUES('$samaccountname','$o365user_serviceName_','$date','Данные добавлены','$o365user_provstat')"
        }
   
    }
    
    }


    #Обновить продукты у пользователя
    if ($UseProductTable_account -ne $null) { 
    
    
    for ($i=0; $i -le $num; $i++) {
    $ServiceName = $o365user.Licenses.ServiceStatus[$i].ServicePlan.ServiceName
    $ProvisioningStatus = $o365user.Licenses.ServiceStatus[$i].ProvisioningStatus
    sql_query "UPDATE UserProductTable SET Date = '$date', Change = 'Данные обновлены', IsActive =  '$ProvisioningStatus' WHERE SAMAccountName like '$samaccountname' and ProductName like '$ServiceName'"  


            }
            }

   } 
}

function Delete_lic_removed_user {

    #----------Delete removed users
    #Запросить всех пользоватлей в БД
    $users_in_TotalTable = sql_query "select SAMAccountName from TotalTable"
    $array_users_in_totaltable = @()
    #$array_users_in_totaltable = $null

    foreach ($user_in_TotalTable in $users_in_TotalTable) {
    $array_users_in_totaltable += $user_in_TotalTable.SAMAccountName + '@_.me'

    }
    $diff_sql_users_and_o365_users = Compare-Object -referenceobject $array_users_in_totaltable -DifferenceObject $o365users.UserPrincipalName | where SideIndicator -eq '<='
    #$diff_sql_users_and_o365_users_names = $diff_sql_users_and_o365_users.InputObject
    #Write-Host "Пользователь был удален из О365: " $diff_sql_users_and_o365_users.InputObject
    #to_log "Пользователь был удален из О365: $diff_sql_users_and_o365_users_names" #убрал, чтобы не забивать лог, при необходимости включить, чтобы видеть отличия в пользователях в БД и О365
    foreach ($diff_sql_user_and_o365_user in $diff_sql_users_and_o365_users) {

    $diff_sql_user_and_o365_user_split = $diff_sql_user_and_o365_user.InputObject.Split("@")
    $diff_sql_user_and_o365_user_sam = $diff_sql_user_and_o365_user_split[0]
    sql_query "UPDATE TotalTable SET LicenseName = NULL, LicenseDate = NULL WHERE SAMAccountName like '$diff_sql_user_and_o365_user_sam'"

    }

    }

function Update_ProductsTable {
    #Get available products
    $products = Get-MsolAccountSku | where AccountSkuId -Like "_:STANDARDWOFFPACK_IW_STUDENT" | Select -ExpandProperty ServiceStatus | Select -ExpandProperty ServicePlan |select ServiceName,ServiceType,TargetClass
    $products_name = $products.ServiceName
    $products_type = $products.ServiceType

#----------Write products to table SQL
    $products_sql = sql_query "select ProductName from ProductsTable"
    $products_sql_name = $products_sql.ProductName
    $Products_diff = Compare-Object -referenceobject $products_sql_name -DifferenceObject $products_name
    $Products_diff_for_remove = Compare-Object -referenceobject $products_sql_name -DifferenceObject $products_name | where SideIndicator -Like "<="
    $Products_diff_name = $Products_diff.InputObject 
    $Products_diff_for_remove_name = $Products_diff_for_remove.InputObject
    foreach ($Product_diff_name in $Products_diff_name) {
    
   $products_ = $products | where ServiceName -Like $Product_diff_name
   $products_name_ = $products_.ServiceType
   sql_query "INSERT INTO ProductsTable (ProductName,Description) VALUES('$Product_diff_name','$products_name_')"

    }

#----------Remove products from table SQL
    foreach ($Product_diff_for_remove_name in $Products_diff_for_remove_name) {
    
   sql_query "delete from ProductsTable where ProductName LIKE $Product_diff_for_remove_name"
   

    }

}

function Update_LicenseProductTable {
$LicenseProductTable_student_lic_products = sql_query "select ProductName from LicenseProductTable where LicenseName LIKE '_:STANDARDWOFFPACK_IW_STUDENT'"
$LicenseProductTable_student_lic_products_name = $LicenseProductTable_student_lic_products.ProductName
$LicenseProductTable_student_lic_products_diff = Compare-Object -referenceobject $LicenseProductTable_student_lic_products_name -DifferenceObject $products_name | where SideIndicator -Like "=>" 
$LicenseProductTable_student_lic_products_diff_name = $LicenseProductTable_student_lic_products_diff.InputObject 
$LicenseProductTable_student_lic_products_diff_remove = Compare-Object -referenceobject $LicenseProductTable_student_lic_products_name -DifferenceObject $products_name | where SideIndicator -Like "<=" 
$LicenseProductTable_student_lic_products_diff_remove_name = $LicenseProductTable_student_lic_products_diff_remove.InputObject

foreach ($LicenseProductTable_student_lic_products_diff_name1 in $LicenseProductTable_student_lic_products_diff_name) {
    
    sql_query "INSERT INTO LicenseProductTable (ProductName,LicenseName) VALUES('$LicenseProductTable_student_lic_products_diff_name1','_:STANDARDWOFFPACK_IW_STUDENT')"
    sql_query "INSERT INTO LicenseProductTable (ProductName,LicenseName) VALUES('$LicenseProductTable_student_lic_products_diff_name1','_:STANDARDWOFFPACK_IW_FACULTY')"
    sql_query "INSERT INTO LicenseProductTable (ProductName,LicenseName) VALUES('$LicenseProductTable_student_lic_products_diff_name1','_:Both')"
    }

    #Remove products
foreach ($LicenseProductTable_student_lic_products_diff_remove_name1 in $LicenseProductTable_student_lic_products_diff_remove_name)
        {
        
        sql_query "delete from LicenseProductTable where ProductName LIKE $LicenseProductTable_student_lic_products_diff_remove_name1"
        
        }


}

Update_ProductsTable #Обновить таблицу ProductsTable
Update_LicenseProductTable #Обновить таблицу LicenseProductTable

#----------Set licenses
try {
$o365users = Get-MsolUser -MaxResults 50000 -ErrorAction Stop
to_log "Debug. Запрос пользователей сервиса О365 - OK" 
}
catch {
to_log "Error. Не удалось запросить пользователей сервиса О365. $_" -ErrorAction Stop

}
$o365users_count = $o365users.Count
to_log "Debug. В О365 $o365users_count пользователей."

Delete_lic_removed_user #Удалить лицензию в базе у несуществующего пользователя О365


foreach ($o365user in $o365users) {
 
#Не учитывать системные и заблокированные учетки
if ($o365user.UserPrincipalName -like "*@_.onmicrosoft.com") #Фильтр слежебных записей.
    {
    
    continue
  
    }

 

if ($o365user.BlockCredential -eq "True") #Фильтр заблокированных пользователей в О365
    {
    
    continue
  
    }


#$date.ToString() + " " + $_.Exception.Message >> C:\tmp\errors.txt
$UPN = $o365user.UserPrincipalName
$samaccountname_split = $UPN.Split("@")
$samaccountname = $samaccountname_split[0]
$user_in_sql = sql_query "select * from TotalTable where SAMAccountName like '$samaccountname'" 
try {
$date_last_sync = $o365user.LastDirSyncTime.AddHours(5).ToString("dd.MM.yyyy HH:mm:ss") 
}
catch {
to_log "Error. У пользователя не установлена дата последней синхронизации с AAD Connect. $_"

}


if ($o365user.PostalCode -eq "Student")
    {
     
    #Write-Host $UPN "User is Student"
    
    #$o365user | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus
    $license_user = $o365user | Select-Object -ExpandProperty Licenses # Получаем свойства поля Licenses
    $AccountSkuId = $license_user.AccountSkuId
    if ($AccountSkuId -like "_:STANDARDWOFFPACK_IW_FACULTY")   #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_FACULTY
        {
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "_:STANDARDWOFFPACK_IW_FACULTY" -ErrorAction "Continue" #Удалить лицензию сотрудника
        to_log "Debug. Пользователь $UPN студент с лицензией _:STANDARDWOFFPACK_IW_FACULTY."
        to_log "Debug. Удаление лицензии сотрудника"
        }
        catch {
        to_log "Error. Не удалось удалить лицензию сотрудника у пользователя $UPN. $_"
        }
        try{
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_STUDENT" -ErrorAction "Continue" #Назначить лицензию студента
        to_log "Debug. Назначение пользователю $UPN лицензии студента"
        }
        catch {
        to_log "Error. Не удалось назначить лицензию студента. $_"
    
        }
        Update_TotalTable_Student_lic 'License changed to _:STANDARDWOFFPACK_IW_STUDENT'

         }
    if ($AccountSkuId -like "_:STANDARDWOFFPACK_IW_STUDENT") #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_STUDENT
         {
     
        to_log "Debug. На студента $UPN уже назначена лицензии студента" 
        Update_TotalTable_Student_lic 'License changed to _:STANDARDWOFFPACK_IW_STUDENT'
     

        }

    
    else #Если не студент и не сотрудник, надо назначить лицензию студента 
        {
        Set-MsolUser -UserPrincipalName $UPN -UsageLocation 'RU' 
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_STUDENT" -ErrorAction "Continue"
        to_log "Debug. На студента $UPN не назначена никакая лицензия. Назначение лицензии студента"
        }
        catch {
        to_log "Error. Ошибка назначения лицензии студента на студента $UPN. $_"
        }
        Update_TotalTable_Student_lic 'License set to _:STANDARDWOFFPACK_IW_STUDENT'
    
        } 
 
        }


if ($o365user.PostalCode -eq "Staff")
    {
    #Write-Host $UPN "User is Staff"
    #$o365user | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus
    $license_user = $o365user | Select-Object -ExpandProperty Licenses # Получаем свойства поля Licenses
    if ($license_user.AccountSkuId -like "_:STANDARDWOFFPACK_IW_STUDENT")  #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_STUDENT
        { 
        to_log "Сотрудник $UPN с лицензией студента."
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "_:STANDARDWOFFPACK_IW_STUDENT" -ErrorAction "Continue" #Удалить лицензию студента
        to-log "Debug. Удаление лицензии студента"
        }
        catch {
        to-log "Error. Ошибка удаления лицензии студента у сотрудника $UPN. $_"
        }
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_FACULTY" -ErrorAction "Continue" #Назначить лицензию сотрудника
        to_log "Debug. Назначение лицензии сотрудника $UPN"
        }
        catch {
        to_log "Error. Ошибка назначения лицензии сотрудника $UPN. $_"
        }
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LO -ErrorAction "Continue" #Назначить продукты О365 с включенным Exchange
        to_log "Debug. Назначение продуктов с включенной почтой сотруднику $UPN"
        }
        catch {
        to_log "Error. Ошибка назначения продуктов с включенной почтой сотруднику $UPN. $_ "
    
        }
    
        Update_TotalTable_Staff_lic 'License changed to _:STANDARDWOFFPACK_IW_FACULTY'
        }
    if ($license_user.AccountSkuId -like "_:STANDARDWOFFPACK_IW_FACULTY") #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_FACULTY
        {
        to_log "Debug. На сотрудника $UPN уже назначена лицензия сотрудника" 
        #Update_TotalTable_Staff_lic 'License changed to _:STANDARDWOFFPACK_IW_FACULTY'
        }
    
    else #Если не студент и не сотрудник, надо назначить лицензию сотрудника с выключенным Exchange
        {
        to_log "Debug. На сотрудника $UPN не назначена никакая лицензия. Назначение лицензии сотрудника"
        Set-MsolUser -UserPrincipalName $UPN -UsageLocation 'RU'
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_FACULTY" -ErrorAction "Continue"
        to_log "Debug. Назначение лицензии сотрудника сотруднику $UPN "
        }
        catch {
        to_log "Error. Ошибка назначения лицензии сотрудника сотруднику $UPN. $_"
        }
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LO_NoExchange -ErrorAction "Continue"
        to_log "Debug. Назначение сотруднику $UPN продуктов О365 за исключением почты"
        }
        catch {
        to_log "Error. Ошибка назначения сотруднику $UPN продуктов О365 за исключением почты. $_ "
    
        }
        Update_TotalTable_Staff_lic 'License set to _:STANDARDWOFFPACK_IW_FACULTY'
        } 
        }

if ($o365user.PostalCode -eq "Both")
    {
    #Write-Host $UPN "User is Both"
    #$get_msoluser | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus
    $license_user = $o365user | Select-Object -ExpandProperty Licenses # Получаем свойства поля Licenses
    if ($license_user.AccountSkuId -like "_:STANDARDWOFFPACK_IW_STUDENT")  #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_STUDENT
        { 
        to_log "Debug. Both $UPN назначена лицензия студента. Назначение лицензии сотрудника с включенной почтой"
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "_:STANDARDWOFFPACK_IW_STUDENT" -ErrorAction "Continue" #Удалить лицензию студента
        to_log "Both $UPN удаление лицензии студента"
        }
        catch 
        {
        to_log "Error. Ошибка удаления лицензии студента у $UPN. $_ "
        }
        #Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_FACULTY" #Назначить лицензию сотрудника
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LO -ErrorAction "Continue" #Назначить продукты О365 с включенным Exchange
        to_log "Debug. Назначить продукты О365 с включенным Exchange $UPN"
        }
        catch {
        to_log "Error. Ошибка назначения продукты О365 с включенным Exchangeу $UPN. $_ "
        
        }
        Update_TotalTable_Staff_lic 'License changed to _:Both'
        }
    if ($license_user.AccountSkuId -like "_:STANDARDWOFFPACK_IW_FACULTY") #Если на пользователя назначена лицензия _:STANDARDWOFFPACK_IW_FACULTY
        {
        to_log "Debug. Пользователь $UPN Both с уже назначенной лицензией сотрудника. Назначение продуктов с включенной почтой"
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LO -ErrorAction "Continue"
        }
        catch {
        to_log "Error. Ошибка назначения продукты О365 с включенным Exchange $UPN. $_ "
        
        } 
        #Update_TotalTable_Both_lic 'License set to _:Both'
        }
    
    else #Если не студент и не сотрудник, надо назначить лицензию сотрудника c включенным Exchange
        {
        to_log "Debug. Пользователь Both $UPN не имеет лицензии. Назначение лицензии в включенной почтой"
        Set-MsolUser -UserPrincipalName $UPN -UsageLocation 'RU'
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "_:STANDARDWOFFPACK_IW_FACULTY" -ErrorAction "Continue"
        to_log "Debug. Назначение лицензии сотрудника пользователю $UPN"
        }
        catch {
        to_log "Error. Ошибка назначения лицензии сотрудника пользователю $UPN. $_"
        
        }
        try {
        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LO -ErrorAction "Continue"
        to_log "Debug. Назначение продуктов с почтой пользователю $UPN"
        }
        catch {
        to_log "Error. Ошибка назначения продуктов с почтой пользователю $UPN. $_"
        }
        Update_TotalTable_Both_lic 'License set to _:Both'
        } 
        }

#$UPN

Update_TotalTable_OfficeSync #Обновляем дату последнего изменения учетной записи в облаке в базе
 
Update_UseProductTable #Обновляем состояние продуктов пользователя
if ($o365user.PostalCode -eq "Error")
        {
        
        to_log "Error. Пользователь $UPN не является ни студентом, ни сотрудником, ни Both. $_"
        
        }
}

#Задание email адреса для ответа - UPN@_.me
to_log "Debug. Задание email адреса для ответа - UPN@_.me пользователям О365, у которых он не задан"
try {
Import-PSSession $Session -AllowClobber -ErrorAction "Stop" 
to_log "Debug. Import-PSSession - OK"
}
catch {
to_log "Error. Import-PSSession. $_"
exit
}
$mail_users = get-mailbox 
foreach ($mail_user in $mail_users) {
if (($mail_user.UserPrincipalName -like "*@_.me") -and ($mail_user.WindowsEmailAddress -notlike "*@_.me") ) {
$UPN_mail = $mail_user.UserPrincipalName
$name_mail = $mail_user.Name
try {
set-mailbox -Identity $name_mail -WindowsEmailAddress $UPN_mail -ErrorAction "Continue"
to_log "Debug. Задание пользователю $UPN_mail атрибута WindowsEmailAddress"
}
catch {
to_log "Error. Ошибка задания пользователю $UPN_mail атрибута WindowsEmailAddress. $_"

}
}

}

Remove-PSSession $Session
to_log "Debug. Работа скрипта завершена"
to_log "------------------------------------"
to_log "------------------------------------"
to_log "------------------------------------"
