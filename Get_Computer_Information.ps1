function Get_PC_Asset_Number{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $Number = [Microsoft.VisualBasic.Interaction]::InputBox("请输入贴在您电脑上的资产编号的数字(eg：HC-PC001，只需输入001)，输入完成后回车")
    $Number = "HC-PC$Number"
    return $Number;
}

function Get_NB_Asset_Number{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $Number = [Microsoft.VisualBasic.Interaction]::InputBox("请输入贴在您笔记本上的资产编号的数字(eg：HC-NB001，只需输入001)，输入完成后回车")
    $Number = "HC-NB$Number"
    return $Number;
}

function Get_User_CN_Name{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $CN_name = [Microsoft.VisualBasic.Interaction]::InputBox("请输入您的中文名字(eg:王德发)，输入完成后回车")
    "中文名字：";
    $CN_name;
}

function Get_User_EN_Name{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $EN_name = [Microsoft.VisualBasic.Interaction]::InputBox("请输入您的英文名字(eg:Nacy Wang)，输入完成后回车")
    "英文名字：";
    $EN_name;
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
    "$Mother_board_facturer $Mother_board_product";
}

function Get_NoteBook_Infotmation{
    $Mother_board_facturer = (Get-WMIObject -Class Win32_Baseboard).Manufacturer;
    $Mother_board_product = (Get-WMIObject -Class Win32_Baseboard).Product;
    "笔记本型号：";
    "$Mother_board_facturer $Mother_board_product";
}

function Get_Mem_Information{
    $Mem_Count = ((Get-WmiObject -Class Win32_PhysicalMemory)| Measure-Object -sum Capacity).count;
    if ($Mem_Count -eq 1){
        $i = $Mem_Count-1;
        $Mem_Space = (Get-WmiObject -Class Win32_PhysicalMemory).Capacity;
        $Men_Manufacturer = (Get-WmiObject -Class Win32_PhysicalMemory).Manufacturer;
        $Mem_Space_GB = "{0:N2}GB" -f (($Mem_Space | Measure-Object -Sum).sum /1gb);
        "内存信息：";
        "内存数量：";
        $Mem_Count;
        "Mem$i";
        $Men_Manufacturer;
        $Mem_Space_GB;
    }
    else{
        "内存信息：";
        "内存数量：";
        $Mem_Count;
        for ($i=0;$i -le $Mem_Count-1;$i++){
            $Mem_Space = (Get-WmiObject -Class Win32_PhysicalMemory)[$i].Capacity;
            $Men_Manufacturer = (Get-WmiObject -Class Win32_PhysicalMemory)[$i].Manufacturer;
            $Mem_Space_GB = "{0:N2}GB" -f (($Mem_Space | Measure-Object -Sum).sum /1gb);
            "Mem$i";
            $Men_Manufacturer;
            $Mem_Space_GB;
        }
    }
}

function Get_Disk_Information{
    $Disk_Count = ((Get-WmiObject -Class Win32_DiskDrive)| Measure-Object -sum Partitions).count;
    if ($Disk_Count -eq 1){
        $i = $Disk_Count-1;
        $Disk_Space = (Get-WmiObject -Class Win32_DiskDrive).Size;
        $Disk_Model = (Get-WmiObject -Class Win32_DiskDrive).Model;
        $Disk_Space_GB = "{0:N2}GB" -f (($Disk_Space | Measure-Object -Sum).sum /1gb);
        "硬盘信息：";
        "硬盘数量：";
        $Disk_Count;
        "Disk$i";
        $Disk_Model;
        $Disk_Space_GB;
    }
    else{
        "硬盘信息：";
        "硬盘数量：";
        $Disk_Count;
        for ($i=0;$i -le $Disk_Count-1;$i++){
            $Disk_Space = (Get-WmiObject -Class Win32_DiskDrive)[$i].Size;
            $Disk_Model = (Get-WmiObject -Class Win32_DiskDrive)[$i].Model;
            $Disk_Space_GB = "{0:N2}GB" -f (($Disk_Space | Measure-Object -Sum).sum /1gb);
            "Disk$i";
            $Disk_Model;
            $Disk_Space_GB;
        }
    }
}

function Get_Network_Adapter_Infotmation{
    $Network_Adapter_Count =((Get-WmiObject -class win32_NetworkAdapterConfiguration -Filter 'ipenabled = "true"').Description| Measure-Object).count;
    $inf = (Get-WmiObject -class win32_NetworkAdapterConfiguration -Filter 'ipenabled = "true"');
    if ($Network_Adapter_Count -eq 0){
        Read-Host "你没有连接到网络,请连接到网络后重新运行此脚本,按任意键退出";
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
    $Gpu_Count = ((Get-WmiObject -class Win32_videoController)| Measure-Object Name).count;
    if ($Gpu_Count -eq 1){
        "显卡信息：";
        "显卡数量：";
        $Gpu_Count
        "Gpu0";
        $Gpu_name = (Get-WmiObject -class Win32_videoController).Name;
        $Gpu_Name;
    }
    else{
        "显卡信息：";
        "显卡数量：";
        $Gpu_Count;
        for($i=0;$i -le $Gpu_Count-1;$i++){    
            $Gpu_Name =(Get-WmiObject -class Win32_videoController)[$i].Name;
            "Gpu$i";
            $Gpu_Name;
        }
    }
}

function Get_Monitor_Infotmation{
    $Monitor_Count = ((Get-WmiObject WmiMonitorID -Namespace root\wmi -Filter 'Active = "true"')| Measure-Object).count;
    if ($Monitor_Count -eq 1){
        "显示器信息：";
        "显示器数量：";
        $Monitor_Count;
        "Monitor0";
        $Monitor_ManufacturerName = ((Get-WmiObject WmiMonitorID -Namespace root\wmi).ManufacturerName | ForEach {[char]$_}) -join "";
        $Monitor_Model = ((Get-WmiObject WmiMonitorID -Namespace root\wmi).UserFriendlyName | ForEach {[char]$_}) -join "";
        $Monitor_ManufacturerName;
        $Monitor_Model;
    }
    else{
        "显示器信息：";
        "显示器数量：";
        $Monitor_Count;
        for($i=0;$i -le $Monitor_Count-1;$i++){    
            $Monitor_ManufacturerName =((Get-WmiObject WmiMonitorID -Namespace root\wmi)[$i].ManufacturerName | ForEach {[char]$_}) -join "";
            $Monitor_Model = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[$i].UserFriendlyName | ForEach {[char]$_}) -join "";
            "Monitor$i";
            $Monitor_ManufacturerName;
            $Monitor_Model;
        }
    }
}

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$pd_computer = [Microsoft.VisualBasic.Interaction]::InputBox("您是在用笔记本电脑吗？是打1回车,否则打0回车")
if ($pd_computer -eq 1)
{
    $title = Get_NB_Asset_Number;
    Get_User_CN_Name |Out-File ./$title'.txt';
    Get_User_EN_Name |Out-File -Append ./$title'.txt';
    Get_Cpu_Information |Out-File  -Append ./$title'.txt';
    Get_NoteBook_Infotmation |Out-File  -Append ./$title'.txt';
    Get_Mem_Information |Out-File  -Append ./$title'.txt';
    Get_Disk_Information |Out-File  -Append ./$title'.txt';
    Get_Network_Adapter_Infotmation |Out-File  -Append ./$title'.txt';
    Get_Gpu_Information |Out-File  -Append ./$title'.txt';
    "电脑信息已获取完成，信息文件已保存在$title.txt中";
    Copy-Item .\$title'.txt' -destination \\192.168.17.40\information\NoteBook;
    "$title.txt已复制到\\192.168.17.40\information\NoteBook中";
    "非常感谢您的配合";
    Read-Host "按任意键退出";
    exit 
}
elseif($pd_computer -eq 0)
{
    $title = Get_PC_Asset_Number;
    Get_User_CN_Name |Out-File ./$title'.txt';
    Get_User_EN_Name |Out-File -Append ./$title'.txt';
    Get_Cpu_Information |Out-File  -Append ./$title'.txt';
    Get_Mother_board_Infotmation |Out-File  -Append ./$title'.txt';
    Get_Mem_Information |Out-File  -Append ./$title'.txt';
    Get_Disk_Information |Out-File  -Append ./$title'.txt';
    Get_Network_Adapter_Infotmation |Out-File  -Append ./$title'.txt';
    Get_Gpu_Information |Out-File  -Append ./$title'.txt';
    Get_Monitor_Infotmation |Out-File  -Append ./$title'.txt';
    "电脑信息已获取完成，信息文件已保存在$title.txt中";
    Copy-Item .\$title'.txt' -destination \\192.168.17.40\information\PC;
    "$title.txt已复制到\\192.168.17.40\information\PC中";
    "非常感谢您的配合";
    Read-Host "按任意键退出";
    exit 
}
else
{
     "输入有误，请重新运行"; 
    Read-Host "按任意键退出";
    exit 
}
