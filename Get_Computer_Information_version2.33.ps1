function Get_User_Name{
    $name = Read-Host "请输入您的名字";
    "用户名字";
    $name;
}

function Get_Asset_Number{
    $Number = Read-Host "请输入贴在您电脑上的资产编号，例如：HC-PC0001";
    return $Number
}
function Get_Cpu_Information{
    $Cpu_name = (Get-WmiObject -class Win32_processor).Name;
    "CPU信息：";
    $Cpu_name;
}

function Get_Mother_board_Infotmation{
    $Mother_board_facturer = (Get-WMIObject -Class Win32_Baseboard).Manufacturer;
    $Mother_board_product = (Get-WMIObject -Class Win32_Baseboard).Product;
    "主板信息：";
    $Mother_board_facturer;
    $Mother_board_product;
}

function Get_Mem_Information{
    $Mem_Count = ((Get-WmiObject -Class Win32_PhysicalMemory)| Measure-Object -sum Capacity).count
    if ($Mem_Count -eq 1){
        $i = $Mem_Count-1;
        $Mem_Space = (Get-WmiObject -Class Win32_PhysicalMemory).Capacity
        $Men_Manufacturer = (Get-WmiObject -Class Win32_PhysicalMemory).Manufacturer 
        $Mem_Space_GB = "{0:N2}GB" -f (($Mem_Space | Measure-Object -Sum).sum /1gb);
        "内存信息：";
        "Mem$i"
        $Men_Manufacturer
        $Mem_Space_GB
    }
    else{
        "内存信息：";
        for ($i=0;$i -le $Mem_Count-1;$i++){
            $Mem_Space = (Get-WmiObject -Class Win32_PhysicalMemory)[$i].Capacity
            $Men_Manufacturer = (Get-WmiObject -Class Win32_PhysicalMemory)[$i].Manufacturer
            $Mem_Space_GB = "{0:N2}GB" -f (($Mem_Space | Measure-Object -Sum).sum /1gb);
            "Mem$i"
            $Men_Manufacturer
            $Mem_Space_GB
        }
    }
}

function Get_Disk_Information{
    $Disk_Count = ((Get-WmiObject -Class Win32_DiskDrive)| Measure-Object -sum Partitions).count
    if ($Disk_Count -eq 1){
        $i = $Disk_Count-1;
        $Disk_Space = (Get-WmiObject -Class Win32_DiskDrive).Size
        $Disk_Model = (Get-WmiObject -Class Win32_DiskDrive).Model 
        $Disk_Space_GB = "{0:N2}GB" -f (($Disk_Space | Measure-Object -Sum).sum /1gb);
        "硬盘信息：";
        "Disk$i"
        $Disk_Model
        $Disk_Space_GB
    }
    else{
        "硬盘信息：";
        for ($i=0;$i -le $Disk_Count-1;$i++){
            $Disk_Space = (Get-WmiObject -Class Win32_DiskDrive)[$i].Size
            $Disk_Model = (Get-WmiObject -Class Win32_DiskDrive)[$i].Model
            $Disk_Space_GB = "{0:N2}GB" -f (($Disk_Space | Measure-Object -Sum).sum /1gb);
            "Disk$i"
            $Disk_Model
            $Disk_Space_GB
        }
    }
}

function Get_Network_Adapter_Infotmation{
    $Network_Adapter_Count =((Get-WmiObject -class win32_NetworkAdapterConfiguration -Filter 'ipenabled = "true"').Description| Measure-Object).count;
    $inf = (Get-WmiObject -class win32_NetworkAdapterConfiguration -Filter 'ipenabled = "true"');
    if ($Network_Adapter_Count -eq 0){
        Read-Host "你没有连接到网络,请连接到网络后重新运行此脚本,按任意键退出" ;
        exit
    }
    elseif ($Network_Adapter_Count -eq 1){
        "网卡MAC：";
        $Mac_Information = $inf.MACAddress;
        $Mac_Information;
        "ip地址：";
        $ip_Information = $inf.ipaddress[0];
        $ip_Information;
    }
    else{
        "网卡MAC：";
        $Mac_Information = $inf.MACAddress[0];
        $Mac_Information;
        "ip地址：";
        $ip_Information = $inf.ipaddress[0];
        $ip_Information;
    }
}

function Get_Gpu_Information{
    $Gpu_name = (Get-WmiObject -class Win32_videoController).Name;
    "显卡信息：";
    $Gpu_name;
}

function Get_Monitor_Infotmation{
    "显示器信息：";
    $A= ((Get-WmiObject WmiMonitorID -Namespace root\wmi).ManufacturerName | ForEach {[char]$_}) -join "" ;
    $A;
    $B= ((Get-WmiObject WmiMonitorID -Namespace root\wmi).UserFriendlyName | ForEach {[char]$_}) -join "";
    $B;
}

$title = Get_Asset_Number
Get_User_Name |Out-File ./$title'.txt';
Get_Cpu_Information |Out-File  -Append ./$title'.txt';
Get_Mother_board_Infotmation |Out-File  -Append ./$title'.txt';
Get_Mem_Information |Out-File  -Append ./$title'.txt';
Get_Disk_Information |Out-File  -Append ./$title'.txt';
Get_Network_Adapter_Infotmation |Out-File  -Append ./$title'.txt';
Get_Gpu_Information |Out-File  -Append ./$title'.txt';
Get_Monitor_Infotmation |Out-File  -Append ./$title'.txt';

"电脑信息已获取完成，信息文件已保存在$title.txt中";
Copy-Item .\$title'.txt' -destination \\192.168.17.40\information
"$title.txt已复制到\\192.168.17.40\information中";
"非常感谢您的配合";
Read-Host "按任意键退出" ;
exit 
