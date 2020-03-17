[CmdletBinding()]
param ( [switch] $reDownload, [int] $threads = 10, [int] $refreshDays = 3)
Begin{
    $downloadScriptBlock = {
        param($url, $outFile, $reDownload)
        
        try{
            $file = ( $url.Substring($url.LastIndexOf("/") + 1) ).Trim()
            
            $platform = ""
            if($file.IndexOf("86") -ne -1 -and $file.IndexOf("64") -eq -1 ){
                $platform = "x86"
            }elseif($file.IndexOf("64") -ne -1 -and $file.IndexOf("86") -eq -1 ){
                $platform = "x64"
            }elseif($file.IndexOf("ia64") -ne -1){
                $platform = "ia64"
            }else{
                $platform = "x86_x64"
            }
            
            if( !(test-path "$outFile\$platform") ){ New-Item -ItemType Directory -Path "$outFile\$platform" -Force}
            
            if( !(test-path "$outFile\$platform\$file") ){
                (new-object system.net.webclient).DownloadFile($url,"$outFile\$platform\$file")
            }elseif($reDownload -eq $true){
                if( test-path "$outFile\$platform\$file"){
                    remove-item "$outFile\$platform\$file"
                }
                (new-object system.net.webclient).DownloadFile($url,"$outFile\$platform\$file")
            }
        }catch [system.exception]{
            return $error 
        }
    }

    class wsusOfflineUpdater{
        [Bool] $reDownload = $false
        [String] $execPath = ""
        [Int] $threads     = 10
        [Int] $refreshDays = 5
        $products  = [ordered]@{
            "win10"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","a3c2375d-0c8a-42f9-bce0-28333e198407", "d2085b71-5f1f-43a9-880d-ed159016d5c6", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94");
            "msse"    = @("6cf036b9-b546-4694-885a-938b93216b66")
            "wd"      = @("8c3fcc84-7410-4a95-8b89-a166a0190486")
            "ofcLive" = @("03c7c488-f8ed-496c-b6e0-be608abb8a79", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "win63"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","405706ed-f1d7-47ea-91e1-eb8860039715", "18e5ea77-e3d1-43b6-a0a8-fa3dbcd42e93", "6407468e-edc7-4ecd-8c32-521f64cee65e", "d31bd4c3-d872-41c9-a2e7-231f372588cb", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4","e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "2c62603e-7a60-4832-9a14-cfdfd2d71b9a");
            "win62"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","2ee2ad83-828c-4405-9479-544d767993fc", "a105a108-7c9b-4518-bbbe-73f0fe30012b", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "0a07aea1-9d09-4c1e-8dc7-7469228d8195");
            "win61"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","bfe5b177-a086-47a0-b102-097e4fa1f807", "f4b9c883-f4db-4fb5-b204-3343c11fa021", "fdfe8200-9d98-44ba-a12a-772282bf60ef", "1556fc1d-f20e-4790-848e-90b7cdbedfda", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94");
            "win60"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","26997d30-08ce-4f25-b2de-699c36a8033a", "ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf", "575d68e2-7c94-48f9-a04f-4b68555d972d", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "6966a762-0c7c-4261-bd07-fb12b4673347", "e9b56b9a-0ca9-4b3e-91d4-bdcf1ac7d94d", "41dce4a6-71dd-4a02-bb36-76984107376d", "ec9aaca2-f868-4f06-b201-fb8eefd84cef", "68623613-134c-4b18-bcec-7497ac1bfcb0" );
            "win52"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e", "dbf57a08-0d5a-46ff-b30c-7715eb9498e9", "032e3af5-1ac5-4205-9ae5-461b4e8cd26d", "a4bedb1d-a809-4f63-9b49-3fe31967b6d0", "4cb6ebd5-e38a-4826-9f76-1416a6f563b0", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "68623613-134c-4b18-bcec-7497ac1bfcb0");
            "win51"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94");
            "win50"   = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "3b4b8621-726e-43a6-b43b-37d07ec7019f");
            "ofc13"   = @("704a0a4a-518f-4d69-9e03-10ba44198bd5", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc10"   = @("84f5f325-30d7-41c4-81d1-87a0e6535b66", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc07"   = @("041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc03"   = @("1403f223-a63f-f572-82ba-c92391218055", "477b856e-65c4-4473-b621-a8b230bb70d9");
        }	
        
        [Array] $badLanguages = @("-af-za_", "-al_", "-am-et_", "-am_", "-arab-iq_", "-arab-pk_", "-as-in_", "-ba_", "-bd_", "-be-by_", "-beta_", "-bg_", "-bgr_", "-bn-bd_", "-bn-in_", "-bn_", "-ca_", "-cat_", "-chs_", "-cht_", "-cs-cz_", "-cs_", "-csy_", "-cym_", "-cyrl-ba_", "-cyrl-tj_", "-da-dk_", "-da_", "-dan_", "-de-de_", "-de_", "-deu_", "-el-gr_", "-el_", "-ell_", "-es-es_", "-es_", "-esn_", "-et_", "-eti_", "-eu_", "-euq_", "-fa-ir_", "-fi-fi_", "-fi_", "-fil-ph_", "-fin_", "-fr-fr_", "-fr_", "-fra_", "-gd-gb_", "-ge_", "-ger_", "-gl_", "-glc_", "-gu-in_", "-hbr_", "-he-il_", "-he_", "-heb_", "-hi_", "-hin_", "-hk_", "-hr_", "-hrv_", "-hu-hu_", "-hun_", "-hy-am_", "-id_", "-ig-ng_", "-in_", "-ind_", "-ir_", "-ire_", "-is-is_", "-is_", "-isl_", "-it-it_", "-it_", "-ita_", "-ja-jp_", "-jpn_", "-ka-ge_", "-ke_", "-kg_", "-kh_", "-km-kh_", "-kn-in_", "-ko-kr_", "-kok-in_", "-kor_", "-ky-kg_", "-latn-ng_", "-latn-uz_", "-lb-lu_", "-lbx_", "-lk_", "-lt_", "-lth_", "-lu_", "-lv_", "-lvi_", "-mi-nz_", "-ml-in_", "-mlt_", "-mn-mn_", "-mn_", "-mr-in_", "-ms-bn_", "-msl_", "-mt-mt_", "-mt_", "-nb-no_", "-nb_", "-ne-np_", "-ng_", "-nl-nl_", "-nl_", "-nld_", "-nn-no_", "-nn_", "-no_", "-non_", "-nor_", "-np_", "-nso-za_", "-nz_", "-or-in_", "-pa-in_", "-pe_", "-ph_", "-pk_", "-pl-pl_", "-pl_", "-plk_", "-pt-br_", "-pt-pt_", "-ptb_", "-ptg_", "-qut-gt_", "-quz-pe_", "-ro-ro_", "-ro_", "-rom_", "-ru-ru_", "-ru_", "-rus_", "-rw-rw_", "-si-lk_", "-sk-sk_", "-sk_", "-sky_", "-sl-si_", "-sl_", "-slv_", "-sq-al_", "-srl_", "-sv-se_", "-sv_", "-sve_", "-sw-ke_", "-ta-in_", "-te-in_", "-tha_", "-ti-et_", "-tk-tm_", "-tm_", "-tn-za_", "-tr-tr_", "-tr_", "-trk_", "-tt-ru_", "-ug-cn_", "-uk-ua_", "-uk_", "-ukr_", "-ur-pk_", "-uz_", "-vit_", "-wo-sn_", "-xh-za_", "-yo-ng_", "-za_", "-zh-cn_", "-zh-hk_", "-zh-tw_", "-zhh_", "-zu-za_", "-af-za_", "-ar-sa_", "-ar_", "-ara_", "-az-latn-", "-bg-bg_", "-bs-latn", "-ca-es", "-cy-gb", "-et-ee", "-eu-es", "-ga-ie", "-gl-es", "-hi-in", "-hr-hr", "-id-id", "-kk-kz", "-lt-lt", "-lv-lv", "-mk-mk", "-ms-my", "-prs-af", "-sr-cyrl", "-sr-latn", "-th-th", "-vi-vn"	)
    
        wsusOfflineUpdater( $reDownload, $threads, $refreshDays ){
            $this.execPath = $PSScriptRoot;
            $this.reDownload = $reDownload
            $this.threads = $threads
            $this.refreshDays = $refreshDays
            
            $this.downloadUpdates()            
            $this.downloadWsusAgent()
            $this.downloadCPP()
            $this.downloadWD()
            $this.downloadMSSE()
            $this.downloadDotNet()
            
            $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            while( $currentThreads -gt 0){
                write-host "Please wait... Awaiting $($currentThreads) threads.  "
                Start-Sleep -seconds 1
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            }
        }
        
        ContainsAny( [string]$s, [string[]]$items ){
            $matchingItems = @($items | where { $s.Contains( $_ ) })
            [bool]$matchingItems
        }
    
        getWebFile($url, $outFile){
            
            $path = (split-path $outFile)
            if( !(test-path $path) ){ New-Item -ItemType Directory -Path $path -Force }
            if( (test-path $outfile) ){ remove-item $outfile }
            
            (new-object system.net.webclient).DownloadFile($url,$outfile)
        }
        
        downloadWsusAgent(){
            write-host "Downloading WSUS Agent"
            if( !( test-path "$($this.execPath)\wsus" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus" }
            if( !( test-path "$($this.execPath)\wsus\wua" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus\wua" }
            
            @(
                "http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/WindowsUpdateAgent30-x86.exe",
                "http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/WindowsUpdateAgent30-x64.exe",
                "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"
            ) | % {
                $file = $_.Substring($_.LastIndexOf("/") + 1)
                
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads){ 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }
                        
                if( !( test-path ".\wsus\wua\$($file)") ){
                    write-host "Downloading $($_)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_ ,(resolve-path "$($this.execPath)\wsus\wua\"),$this.reDownload)
                    
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\wua\$file").LastWriteTime (get-date) ).Days -gt 3){
                        write-host "Downloading $($_)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_ ,(resolve-path "$($this.execPath)\wsus\wua\"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $file"
                    }
                }
            }
            
            if(test-path "$($this.execPath)\wsus\wua\package.cab"){ remove-item "$($this.execPath)\wsus\wua\package.cab" }
            if(test-path "$($this.execPath)\wsus\wua\package.xml"){ remove-item "$($this.execPath)\wsus\wua\package.xml" }
            invoke-expression("expand.exe '$($this.execPath)\wsus\wua\wsusscn2.cab' -F:package.cab '$($this.execPath)\wsus\wua\'")
            invoke-expression("expand.exe '$($this.execPath)\wsus\wua\package.cab' '$($this.execPath)\wsus\wua\package.xml'")
            
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-host "Finished Downloading WSUS Agent"
        }
        
        downloadWD(){
            write-host "Downloading Windows Defender"
            if( !( test-path "$($this.execPath)\wsus" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus" }
            if( !( test-path "$($this.execPath)\wsus\wd\" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus\wd" }
            
            @(
                "http://download.microsoft.com/download/DefinitionUpdates/mpas-feX64.exe,mpas-feX64.exe",
                "http://download.microsoft.com/download/DefinitionUpdates/mpas-fe.exe,mpas-fe.exe"
            ) | % {
                $url = $_.Substring(0,$_.LastIndexOf(","))
                $file = $_.Substring($_.LastIndexOf(",") + 1)
                
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads){ 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }
                            
                            
                if( !( test-path "$($this.execPath)\wsus\wd\$($file)") ){
                    write-host "Downloading $($url)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\wd\"),$this.reDownload)
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\wd\$file").LastWriteTime (get-date) ).Days -gt $this.RefreshDays){
                        write-host "Downloading $($url)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\wd\"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $($file)"
                    }
                }
            }
            
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-host "Finished Downloading Windows Defender"
        }
    
        downloadMSSE(){
            write-host "Downloading Microsoft Security Essentials"
            if( !( test-path "$($this.execPath)\wsus" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus" }
            
            if( !( test-path "$($this.execPath)\wsus\msse\" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus\msse" }
            
            @(
                "http://download.microsoft.com/download/A/3/8/A38FFBF2-1122-48B4-AF60-E44F6DC28BD8/ENUS/amd64/MSEInstall.exe,MSEInstall-x64-enu.exe",
                "http://download.microsoft.com/download/DefinitionUpdates/mpam-fex64.exe,mpam-fex64.exe",
                "http://definitionupdates.microsoft.com/download/DefinitionUpdates/NRI/amd64/nis_full.exe,nis_full_x64.exe"
            ) | % {
                $url = $_.Substring(0,$_.LastIndexOf(","))
                $file = $_.Substring($_.LastIndexOf(",") + 1)
                
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads){ 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }

                if( !( test-path ".\wsus\msse\$($file)") ){
                    write-host "Downloading $($url)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\msse\"),$this.reDownload)
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\msse\$file").LastWriteTime (get-date) ).Days -gt $this.RefreshDays){
                        write-host "Downloading $($url)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\msse\"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $($file)"
                    }
                }
            }
            
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-host "Finished Downloading Microsoft Security Essentials"
        }
    
        downloadCPP(){
            write-host "Downloading C++ Redistributables"
            
            if( !( test-path ".\wsus" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus" }
            if( !( test-path ".\wsus\cpp\" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus\cpp" }
            
            
            @(
                "https://aka.ms/vs/16/release/vc_redist.x86.exe",
                "https://aka.ms/vs/16/release/vc_redist.x64.exe",
                "https://aka.ms/vs/16/release/VC_redist.arm64.exe"
            ) | % {
            
                $url = $_
                $file = $_.Substring($_.LastIndexOf("/") + 1)
                
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads ) { 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }

                if( !( test-path ".\wsus\cpp\$($file)") ){
                    write-host "Downloading $($url)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\cpp\"),$this.reDownload)
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\cpp\$file").LastWriteTime (get-date) ).Days -gt $this.RefreshDays){
                        write-host "Downloading $($url)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\cpp\"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $($file)"
                    }
                }
            }
            
            
            [xml] $xml = get-content "$($this.execPath)\wsus\wua\package.xml"
            $ns = new-object Xml.XmlNamespaceManager $xml.NameTable
            $ns.AddNamespace('dns', 'http://schemas.microsoft.com/msus/2004/02/OfflineSync')
            
            $nodes = $xml.selectnodes("//dns:FileLocation[contains(@Url, 'vcredist')]",$ns)
            $nodes | % {
                $url = $_.Url
                $file = $_.Url.Substring($_.Url.LastIndexOf("/") + 1)
                
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads){ 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }

                if( !( test-path ".\wsus\cpp\$($file)") ){
                    write-host "Downloading $($url)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\cpp\"),$this.reDownload)
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\cpp\$file").LastWriteTime (get-date) ).Days -gt $this.RefreshDays){
                        write-host "Downloading $($url)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($url ,(resolve-path "$($this.execPath)\wsus\cpp\"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $($file)"
                    }
                }
            }
            
            
            
            
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-host "Finished Downloading C++ Redistributables"
        }
    
        downloadDotNet(){
            write-host "Downloading DotNet Files"
            if( !( test-path "$($this.execPath)\wsus" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus" }
            if( !( test-path ".\wsus\dotNet\" ) ){ New-Item -ItemType directory -Path "$($this.execPath)\wsus\dotNet" }
            
            @(
                "http://download.microsoft.com/download/2/0/e/20e90413-712f-438c-988e-fdaa79a8ac3d/dotnetfx35.exe",
                "http://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe",
                "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe",
                "https://download.microsoft.com/download/F/9/4/F942F07D-F26F-4F30-B4E3-EBD54FABA377/NDP462-KB3151800-x86-x64-AllOS-ENU.exe",
                "https://download.microsoft.com/download/6/E/4/6E48E8AB-DC00-419E-9704-06DD46E5F81D/NDP472-KB4054530-x86-x64-AllOS-ENU.exe",
                "https://download.visualstudio.microsoft.com/download/pr/014120d7-d689-4305-befd-3cb711108212/0fd66638cde16859462a6243a4629a50/ndp48-x86-x64-allos-enu.exe"

            ) | % {
                $file = $_.Substring($_.LastIndexOf("/") + 1)
                        
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                While ( $currentThreads -ge $this.threads){ 
                    write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                    Start-Sleep -seconds 1
                    $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                }


                if( !( test-path "$($this.execPath)\wsus\dotNet\$($file)") ){
                    write-host "Downloading $($file)"
                    Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_ , (resolve-path "$($this.execPath)\wsus\dotNet"),$this.reDownload)
                }else{
                    if( (new-timespan (get-item "$($this.execPath)\wsus\dotNet\$($file)").LastWriteTime (get-date) ).Days -gt $this.RefreshDays){
                        write "Re-Downloading $($file)"
                        Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_ ,(resolve-path "$($this.execPath)\wsus\dotNet"),$this.reDownload)
                    }else{
                        write-host "Using pre-existing $($file)"
                    }
                }
            }
            
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-host "Finished Downloading DotNet Files"
        }
        
        downloadUpdates(){
            write-host "Downloading Windows Updates"
            [xml] $xml = get-content "$($this.execPath)\wsus\wua\package.xml"
            $ns = new-object Xml.XmlNamespaceManager $xml.NameTable
            $ns.AddNamespace('dns', 'http://schemas.microsoft.com/msus/2004/02/OfflineSync')
            
            $productIndex = 0
            
            foreach($product in ( $this.products.keys  ) ){
                write-host "Downloading Updates for $($product)"
                @("","x86","x64","x86_x64","ia64") | % {
                    if( !(test-path "$($this.execPath)\wsus\$product\$($_)" ) ){
                        New-Item -ErrorAction SilentlyContinue -ItemType directory -Path "$($this.execPath)\wsus\$product\$($_)" 
                    }
                }
                
                foreach($prodId in $this.products.$product){
                    
                    $productNodes = $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '$($prodId)']]",$ns)
                    if($productNodes -ne $null -and $productNodes.count -gt 1) {
                    
                        $productNodes | %{
                            
                            $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                            While ( $currentThreads -ge $this.threads){ 
                                write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                                Start-Sleep -seconds 1
                                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                            }
                            
                            if($_.EulaFiles -ne $null){
                                $xml.SelectNodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[@RevisionId='$($_.RevisionId)']/dns:EulaFiles/dns:File[./dns:Language/@Name='en']/@Id",$ns) | %{
                                    $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$($_.'#text')']",$ns) | % {
                                        if( ! $this.ContainsAny($_.Url,$this.badLanguages) ){
                                            write-host "Downloading file $($_.Url.substring($_.url.lastindexof('/')+1)) for $product"
                                            Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_.Url ,(resolve-path "$($this.execPath)\wsus\$product\"),$this.reDownload)
                                        }
                                    }
                                }
                            }
                             
                            $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                            While ( $currentThreads -ge $this.threads){ 
                                write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                                Start-Sleep -seconds 1
                                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
                            }
                            $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='$($_.RevisionId)'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id",$ns) | % {
                                $fileId = $($_.'#text')
                                
                                $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$($_.'#text')']",$ns) | % {
                                    if(!$this.ContainsAny($_.Url,$this.badLanguages)){
                                        if($_.Url -notLike '*mui*'){
                                            write-host "Downloading file $($_.Url.substring($_.url.lastindexof('/')+1)) for $product"
                                            Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_.Url ,(resolve-path "$($this.execPath)\wsus\$product\"),$this.reDownload)
                                        }
                                    }
                                }
                            }
                            
                            get-job | ? { $_.State -eq 'Completed' } | remove-job
                            [GC]::Collect()
                        }
                    }
                    
                    $deleteNodes = $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[./dns:SupersededBy][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '$($prodId)']]",$ns)
                    if($deleteNodes -ne $null -and $deleteNodes.count -gt 1){
                        $deleteNodes | % {
                            
                            $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='$($_.RevisionId)'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id",$ns) | % {
                                $fileId = $($_.'#text')
                                $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$($_.'#text')']",$ns) | % {
                                    $fileName = $($_.Url.substring($_.url.lastindexof('/')+1))
                                    if( (test-path "$($this.execPath)\wsus\$product\x86\$filename" ) )    { write-host ("Deleting Superseded file $($fileName) for $product"); remove-item "$($this.execPath)\wsus\$product\x86\$filename" }
                                    if( (test-path "$($this.execPath)\wsus\$product\x64\$filename" ) )    { write-host ("Deleting Superseded file $($fileName) for $product"); remove-item "$($this.execPath)\wsus\$product\x64\$filename" }
                                    if( (test-path "$($this.execPath)\wsus\$product\ia64\$filename" ) )   { write-host ("Deleting Superseded file $($fileName) for $product"); remove-item "$($this.execPath)\wsus\$product\ia64\$filename" }
                                    if( (test-path "$($this.execPath)\wsus\$product\x86_x64\$filename" ) ){ write-host ("Deleting Superseded file $($fileName) for $product"); remove-item "$($this.execPath)\wsus\$product\x86_x64\$filename" }
                                }
                            }
                        }
                    }
                }
            }
            
            sleep -seconds 10
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            [GC]::Collect()
            write-host "Finished Downloading Windows Updates"
        }
    
    }
}
Process{
    $wsus = [wsusOfflineUpdater]::new($reDownload, $threads, $refreshDays)
}
End{

}
