[CmdletBinding()]
param ( [switch] $reDownload, [int] $threads = 10, [int] $refreshDays = 3)
Begin{
    $downloadScriptBlock = {
        param($url, $outfile, $reDownload)

        try{
            $download = $false

            $path = split-path $outfile
            if( !(test-path "$path") ){ New-Item -ItemType Directory -Path "$path" -Force}

            #file does not exist
            if( !(test-path "$outfile") ){
                $download = $true
            }

            #file exists
            if( test-path "$outfile"){
                #hash mismatch
                if( (get-item -path "$outfile" | select -expand name) -notlike "*" + (Get-FileHash -Algorithm sha1 "$outfile" | select -expand hash ) + "*"){
                    $download = $true
                }

                #force download on
                if($reDownload -eq $true){ #forced redownload
                    $download = $true
                }
            }

            if($download){
                write-host "Downloading $outfile"
                if( test-path "$outfile"){
                    remove-item "$outFile"
                }

                (new-object system.net.webclient).DownloadFile($url,"$outfile")
            }else{
                write-host "Using existing $outfile"
            }
        }catch [system.exception]{
            write-error $error
            return $error
        }
    }

    class wsusOfflineUpdater{
        [Bool] $reDownload = $false
        [String] $execPath = ""
        [Int] $threads     = 10
        [Int] $refreshDays = 3
        $products  = [ordered]@{
            "sql2k"    = @("7145181b-9556-4b11-b659-0162fa9df11f", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "c96c35fc-a21f-481b-917c-10c4f64792cb");
            "sql2k5"   = @("60916385-7546-4e9b-836e-79d65e517bab", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "c96c35fc-a21f-481b-917c-10c4f64792cb");
            "sqlsk8"   = @("dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "c5f0b23c-e990-4b71-9808-718d353f533a", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "c96c35fc-a21f-481b-917c-10c4f64792cb");
            "sql2k8r2" = @("bb7bc3a7-857b-49d4-8879-b639cf5e8c3c", "e9ece729-676d-4b57-b4d1-7e0ab0589707", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "c96c35fc-a21f-481b-917c-10c4f64792cb", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a");
            "sql2k12"  = @("56750722-19b4-4449-a547-5b68f19eee38", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "c96c35fc-a21f-481b-917c-10c4f64792cb", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "7fe4630a-0330-4b01-a5e6-a77c7ad34eb0");
            "sql2k14"  = @("caab596c-64f2-4aa9-bbe3-784c6e2ccf9c", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "c96c35fc-a21f-481b-917c-10c4f64792cb", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "892c0584-8b03-428f-9a74-224fcd6887c0");
            "sql2k16"  = @("93f0b0bc-9c20-4ca5-b630-06eb4706a447", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "c96c35fc-a21f-481b-917c-10c4f64792cb", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "892c0584-8b03-428f-9a74-224fcd6887c0");
            "sql2k17"  = @("ca6616aa-6310-4c2d-a6bf-cae700b85e86", "dee854fd-e9d2-43fd-bbc3-f7568e3ce324", "fe324c6a-dac1-aca8-9916-db718e48fa3a", "c96c35fc-a21f-481b-917c-10c4f64792cb", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e");

            "ofcLive"  = @("03c7c488-f8ed-496c-b6e0-be608abb8a79", "477b856e-65c4-4473-b621-a8b230bb70d9", "ec231084-85c2-4daf-bfc4-50bbe4022257");
            "ofc365"   = @("30eb551c-6288-4716-9a78-f300ec36d72b", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc16"    = @("25aed893-7c2d-4a31-ae22-28ff8ac150ed", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc13"    = @("704a0a4a-518f-4d69-9e03-10ba44198bd5", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc10"    = @("84f5f325-30d7-41c4-81d1-87a0e6535b66", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc07"    = @("041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofc03"    = @("1403f223-a63f-f572-82ba-c92391218055", "477b856e-65c4-4473-b621-a8b230bb70d9");
            "ofcxp"    = @("6248b8b1-ffeb-dbd9-887a-2acf53b09dfe", "477b856e-65c4-4473-b621-a8b230bb70d9")

            "win10"    = @("569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5", "3c54bb6c-66d1-4a79-884c-8a0c96fa20d1", "6964aab4-c5b5-43bd-a17d-ffb4346a8e1d", "a3c2375d-0c8a-42f9-bce0-28333e198407", "d2085b71-5f1f-43a9-880d-ed159016d5c6", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "05eebf61-148b-43cf-80da-1c99ab0b8699", "34f268b4-7e2d-40e1-8966-8bb6ea3dad27", "bab879a4-c1af-4b52-9617-0f9ae1286fb6", "0ba562e6-a6ba-490d-bdce-93a770ba8d21", "cfe7182c-14a0-4d7e-9f5e-505d5c3a66f6", "f5b5092c-d05e-4eb1-8a6a-919770378ff6", "06da2f0c-7937-4e28-b46c-a37317eade73", "e4b04398-adbd-4b69-93b9-477322331cd3", "876ad18f-f41d-442a-ac64-f5c5ce74cc83", "c70f1038-66ac-443d-9e58-ac22e891e4fb", "e104dd76-2895-41c4-9eb5-c483a61e9427", "3efabf46-3037-4c85-a752-3189e574b621", "6111a83d-7a6b-4a2c-a7c2-f222eebcabf4", "abc45868-0c9c-4bc0-a36d-03d54113baf4", "7d247b99-caa2-45e4-9c8f-6d60d0aae35c", "fc7c9913-7a1e-4b30-b602-3c62fffd9b1a", "d2085b71-5f1f-43a9-880d-ed159016d5c6", "c1006636-eab4-4b0b-b1b0-d50282c0377e", "bb06ba08-3df8-4221-8794-18effb79156a", "b7f52cfb-c9e9-4481-9bc0-c8b4e208ba39", "a3c2375d-0c8a-42f9-bce0-28333e198407");
            "win63"    = @("d31bd4c3-d872-41c9-a2e7-231f372588cb", "8b4e84f6-595f-41ed-854f-4ca886e317a5", "bfd3e48c-c96b-43fd-8b09-98cdc89dc77e", "f3c2263d-b256-4c49-a246-973c0e366449", "01030579-66d2-446e-8c65-538df07e0e44", "14a011c7-d17b-4b71-a2a4-051807f4f4c6", "f7b29b7a-086b-43f9-9cc8-e1a2f8a31e08", "6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","405706ed-f1d7-47ea-91e1-eb8860039715", "18e5ea77-e3d1-43b6-a0a8-fa3dbcd42e93", "6407468e-edc7-4ecd-8c32-521f64cee65e", "d31bd4c3-d872-41c9-a2e7-231f372588cb", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4","e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "2c62603e-7a60-4832-9a14-cfdfd2d71b9a");
            "win62"    = @("26cbba0f-45de-40d5-b94a-3cbe5b761c9d", "97c4cee8-b2ae-4c43-a5ee-08367dab8796", "3e5cc385-f312-4fff-bd5e-b88dcf29b476", "589db546-7849-47f5-bbc0-1f66cf12f5c2", "393789f5-61c1-4881-b5e7-c47bcca90f94", "6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","2ee2ad83-828c-4405-9479-544d767993fc", "a105a108-7c9b-4518-bbbe-73f0fe30012b", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "0a07aea1-9d09-4c1e-8dc7-7469228d8195");
            "win61"    = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","bfe5b177-a086-47a0-b102-097e4fa1f807", "f4b9c883-f4db-4fb5-b204-3343c11fa021", "fdfe8200-9d98-44ba-a12a-772282bf60ef", "1556fc1d-f20e-4790-848e-90b7cdbedfda", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94");
            "win60"    = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","26997d30-08ce-4f25-b2de-699c36a8033a", "ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf", "575d68e2-7c94-48f9-a04f-4b68555d972d", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "6966a762-0c7c-4261-bd07-fb12b4673347", "e9b56b9a-0ca9-4b3e-91d4-bdcf1ac7d94d", "41dce4a6-71dd-4a02-bb36-76984107376d", "ec9aaca2-f868-4f06-b201-fb8eefd84cef", "68623613-134c-4b18-bcec-7497ac1bfcb0" );
            "win52"    = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e", "dbf57a08-0d5a-46ff-b30c-7715eb9498e9", "032e3af5-1ac5-4205-9ae5-461b4e8cd26d", "a4bedb1d-a809-4f63-9b49-3fe31967b6d0", "4cb6ebd5-e38a-4826-9f76-1416a6f563b0", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "68623613-134c-4b18-bcec-7497ac1bfcb0");
            "win51"    = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94");
            "win50"    = @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "3b4b8621-726e-43a6-b43b-37d07ec7019f");
        }

        [Array] $badLanguages = @("-af-za_", "-al_", "-am-et_", "-am_", "-arab-iq_", "-arab-pk_", "-as-in_", "-ba_", "-bd_", "-be-by_", "-beta_", "-bg_", "-bgr_", "-bn-bd_", "-bn-in_", "-bn_", "-ca_", "-cat_", "-chs_", "-cht_", "-cs-cz_", "-cs_", "-csy_", "-cym_", "-cyrl-ba_", "-cyrl-tj_", "-da-dk_", "-da_", "-dan_", "-de-de_", "-de_", "-deu_", "-el-gr_", "-el_", "-ell_", "-es-es_", "-es_", "-esn_", "-et_", "-eti_", "-eu_", "-euq_", "-fa-ir_", "-fi-fi_", "-fi_", "-fil-ph_", "-fin_", "-fr-fr_", "-fr_", "-fra_", "-gd-gb_", "-ge_", "-ger_", "-gl_", "-glc_", "-gu-in_", "-hbr_", "-he-il_", "-he_", "-heb_", "-hi_", "-hin_", "-hk_", "-hr_", "-hrv_", "-hu-hu_", "-hun_", "-hy-am_", "-id_", "-ig-ng_", "-in_", "-ind_", "-ir_", "-ire_", "-is-is_", "-is_", "-isl_", "-it-it_", "-it_", "-ita_", "-ja-jp_", "-jpn_", "-ka-ge_", "-ke_", "-kg_", "-kh_", "-km-kh_", "-kn-in_", "-ko-kr_", "-kok-in_", "-kor_", "-ky-kg_", "-latn-ng_", "-latn-uz_", "-lb-lu_", "-lbx_", "-lk_", "-lt_", "-lth_", "-lu_", "-lv_", "-lvi_", "-mi-nz_", "-ml-in_", "-mlt_", "-mn-mn_", "-mn_", "-mr-in_", "-ms-bn_", "-msl_", "-mt-mt_", "-mt_", "-nb-no_", "-nb_", "-ne-np_", "-ng_", "-nl-nl_", "-nl_", "-nld_", "-nn-no_", "-nn_", "-no_", "-non_", "-nor_", "-np_", "-nso-za_", "-nz_", "-or-in_", "-pa-in_", "-pe_", "-ph_", "-pk_", "-pl-pl_", "-pl_", "-plk_", "-pt-br_", "-pt-pt_", "-ptb_", "-ptg_", "-qut-gt_", "-quz-pe_", "-ro-ro_", "-ro_", "-rom_", "-ru-ru_", "-ru_", "-rus_", "-rw-rw_", "-si-lk_", "-sk-sk_", "-sk_", "-sky_", "-sl-si_", "-sl_", "-slv_", "-sq-al_", "-srl_", "-sv-se_", "-sv_", "-sve_", "-sw-ke_", "-ta-in_", "-te-in_", "-tha_", "-ti-et_", "-tk-tm_", "-tm_", "-tn-za_", "-tr-tr_", "-tr_", "-trk_", "-tt-ru_", "-ug-cn_", "-uk-ua_", "-uk_", "-ukr_", "-ur-pk_", "-uz_", "-vit_", "-wo-sn_", "-xh-za_", "-yo-ng_", "-za_", "-zh-cn_", "-zh-hk_", "-zh-tw_", "-zhh_", "-zu-za_", "-af-za_", "-ar-sa_", "-ar_", "-ara_", "-az-latn-", "-bg-bg_", "-bs-latn", "-ca-es", "-cy-gb", "-et-ee", "-eu-es", "-ga-ie", "-gl-es", "-hi-in", "-hr-hr", "-id-id", "-kk-kz", "-lt-lt", "-lv-lv", "-mk-mk", "-ms-my", "-prs-af", "-sr-cyrl", "-sr-latn", "-th-th", "-vi-vn"	)

        wsusOfflineUpdater( $reDownload, $threads, $refreshDays ){
            $this.execPath = $PSScriptRoot;
            $this.reDownload = $reDownload
            $this.threads = $threads
            $this.refreshDays = $refreshDays

            #$this.downloadWsusAgent()
            #$this.downloadWd()
            #$this.downloadCppDotNet()

            $this.downloadUpdates()

            $this.awaitThreads(0)
        }

        ContainsAny( [string]$s, [string[]]$items ){
            $matchingItems = @($items | where { $s.Contains( $_ ) })
            [bool]$matchingItems
        }

        logger($level, $msg){
            $message = "$( Get-Date -Format 'yyyy-MM-dd HH:mm:ss,fff' ) - microsoft_updates - $($level) - $($msg)"
            add-content -Path "$($this.execPath)\microsoft_updates.log" -value $message
            write-host $message
        }
        
        awaitThreads($count){
            $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            While ( $currentThreads -gt $count){
                write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                Start-Sleep -seconds 5
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            }
        }

        awaitThreads(){
            $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            While ( $currentThreads -ge $this.threads){
                write-host "Please wait... Currently using $($currentThreads) of $($this.threads) threads.  "
                Start-Sleep -seconds 5
                $currentThreads = (Get-Job | ? {$_.State -eq "Running"}).Count
            }
        }

        downloadWsusAgent(){
            $this.logger("INFO", "Downloading WSUS Agent")

            @(
                @('http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab', "$($this.execPath)/storage/wua/wsusscn2.cab"),
                @('http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/windowsupdateagent30-x86.exe', "$($this.execPath)/storage/wua/x86/windowsupdateagent30-x86.exe"),
                @('http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/windowsupdateagent30-x64.exe', "$($this.execPath)/storage/wua/x64/windowsupdateagent30-x64.exe")
            ) | % {
                $this.awaitThreads()
                Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_[0] ,$_[1],$this.reDownload)
            }
            $this.awaitThreads(0)

            if(test-path "$($this.execPath)\storage\wua\package.cab"){
                $this.logger("WARNING", "Removing old package.cab")
                remove-item "$($this.execPath)\storage\wua\package.cab"
            }

            if(test-path "$($this.execPath)\storage\wua\package.xml"){
                $this.logger("WARNING", "removing old package.xml")
                remove-item "$($this.execPath)\storage\wua\package.xml"
            }

            invoke-expression("expand.exe '$($this.execPath)\storage\wua\wsusscn2.cab' -F:package.cab '$($this.execPath)\storage\wua\'")
            invoke-expression("expand.exe '$($this.execPath)\storage\wua\package.cab' '$($this.execPath)\storage\wua\package.xml'")

            get-job | ? { $_.State -eq 'Completed' } | remove-job
            write-$this.logger("INFO", "Finished Downloading WSUS Agent")
        }

        downloadWd(){
            $this.logger("INFO", "Downloading Windows Defender")

            @(
                @('https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64', "$($this.execPath)/storage/wd/x64/mpam-feX64.exe"),
                @('https://go.microsoft.com/fwlink/?LinkID=121721&arch=x86', "$($this.execPath)/storage/wd/x86/mpam-feX86.exe")
            ) | % {
                $this.awaitThreads()
                Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_[0] ,$_[1],$this.reDownload)
            }
            $this.awaitThreads(0)

            get-job | ? { $_.State -eq 'Completed' } | remove-job
            $this.logger("INFO", "Finished Downloading Windows Defender" )
        }

        downloadCppDotNet(){
            $this.logger("INFO", "Downloading C++ Redistributables")

            @(
                @('http://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.EXE', "$($this.execPath)/storage/cpp_dn/x86/vcredist2005_x86.exe"),
                @('http://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe', "$($this.execPath)/storage/cpp_dn/x86/vcredist2008_x86.exe"),
                @('http://download.microsoft.com/download/E/E/0/EE05C9EF-A661-4D9E-BCE2-6961ECDF087F/vcredist_x86.exe', "$($this.execPath)/storage/cpp_dn/x86/vcredist2010_x86.exe"),
                @('http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe', "$($this.execPath)/storage/cpp_dn/x86/vcredist2012_x86.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/10912113/5da66ddebb0ad32ebd4b922fd82e8e25/vcredist_x86.exe', "$($this.execPath)/storage/cpp_dn/x86/vcredist2013_x86.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/9307e627-aaac-42cb-a32a-a39e166ee8cb/E59AE3E886BD4571A811FE31A47959AE5C40D87C583F786816C60440252CD7EC/VC_redist.x86.exe', "$($this.execPath)/storage/cpp_dn/x86/vcredist2019_x86.exe"),

                @('http://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE', "$($this.execPath)/storage/cpp_dn/x64/vcredist2005_x64.exe"),
                @('http://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe', "$($this.execPath)/storage/cpp_dn/x64/vcredist2008_x64.exe"),
                @('http://download.microsoft.com/download/E/E/0/EE05C9EF-A661-4D9E-BCE2-6961ECDF087F/vcredist_x64.exe', "$($this.execPath)/storage/cpp_dn/x64/vcredist2010_x64.exe"),
                @('http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe', "$($this.execPath)/storage/cpp_dn/x64/vcredist2012_x64.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/10912041/cee5d6bca2ddbcd039da727bf4acb48a/vcredist_x64.exe', "$($this.execPath)/storage/cpp_dn/x64/vcredist2013_x64.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/3b070396-b7fb-4eee-aa8b-102a23c3e4f4/40EA2955391C9EAE3E35619C4C24B5AAF3D17AEAA6D09424EE9672AA9372AEED/VC_redist.x64.exe', "$($this.execPath)/storage/cpp_dn/x64/vcredist2019_x64.exe"),

                @('http://download.windowsupdate.com/c/msdownload/update/software/updt/2017/10/ndp46-kb4041778-x86_7eb6b6850959824025c3d3e217391bcb512b30c7.exe', "$($this.execPath)/storage/cpp_dn/x86/ndp46-kb4041778-x86.exe"),
                @('http://download.windowsupdate.com/c/msdownload/update/software/updt/2017/10/ndp46-kb4041778-x64_1e36c362f9f6c4dfdef1a3c0abe4fb560ec7f25c.exe', "$($this.execPath)/storage/cpp_dn/x64/ndp46-kb4041778-x64.exe"),
                @('http://download.microsoft.com/download/2/0/e/20e90413-712f-438c-988e-fdaa79a8ac3d/dotnetfx35.exe', "$($this.execPath)/storage/cpp_dn/x86_x64/dotnetfx35.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/014120d7-d689-4305-befd-3cb711108212/0fd66638cde16859462a6243a4629a50/ndp48-x86-x64-allos-enu.exe', "$($this.execPath)/storage/cpp_dn/x64/ndp48-x86-x64-allos-enu.exe"),
                @('http://download.visualstudio.microsoft.com/download/pr/7afca223-55d2-470a-8edc-6a1739ae3252/65592cace88fcfb3e14a5c4a54833794/ndp48-x86-x64-allos-deu.exe',"$($this.execPath)/storage/cpp_dn/x86/ndp48-x86-x86-allos-enu.exe")
            ) | % {
                $this.awaitThreads()
                Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_[0] ,$_[1],$this.reDownload)

            }

            $this.awaitThreads(0)
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            $this.logger("INFO", "Finished Downloading C++ Redistributables")
        }

        downloadUpdates(){
            $this.logger("INFO", "Downloading Windows Updates")

            [xml] $xml = get-content "$($this.execPath)\storage\wua\package.xml"
            $ns = new-object Xml.XmlNamespaceManager $xml.NameTable
            $ns.AddNamespace('dns', 'http://schemas.microsoft.com/msus/2004/02/OfflineSync')
            $productIndex = 0

            foreach($product in ( $this.products.keys  ) ){
                $this.logger("INFO", "Downloading Updates for $($product)")

                foreach($prodId in $this.products.$product){
                    $this.logger("INFO", "Sub-Product $($prodId)")
                    
                    $productNodes = $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '$($prodId)']]",$ns)
                    if($productNodes -ne $null -and $productNodes.count -gt 1) {
                        $productNodes | %{
                            $this.awaitThreads()

                            if($_.EulaFiles -ne $null){
                                $xml.SelectNodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[@RevisionId='$($_.RevisionId)']/dns:EulaFiles/dns:File[./dns:Language/@Name='en']/@Id",$ns) | %{
                                    $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$($_.'#text')']",$ns) | % {
                                        if( ! $this.ContainsAny($_.Url,$this.badLanguages) ){
                                            $arch = "x86_x64"
                                            if( $_.Url -like '*arm*'){
                                                $arch = "arm64"
                                            }elseif( $_.Url -like '*ia64*'){
                                                $arch = "ia64"
                                            }elseif( $_.Url -like '*x86*'){
                                                $arch = "x86"
                                            }elseif( $_.Url -like '*x64*'){
                                                $arch = "x64"
                                            }
                                            $this.awaitThreads()
                                            $filename = (split-path $_.Url -leaf)
                                            $this.logger("INFO", "Checking $filename")
                                            Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_.Url , "$($this.execPath)\storage\$product\$arch\$filename", $this.reDownload)
                                        }
                                    }
                                }
                            }

                            $this.awaitThreads()

                            $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='$($_.RevisionId)'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id",$ns) | % {
                                $fileId = $($_.'#text')

                                $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$($_.'#text')']",$ns) | % {
                                    if(!$this.ContainsAny($_.Url,$this.badLanguages)){
                                        if($_.Url -notLike '*mui*'){

                                            $arch = "x86_x64"
                                            if( $_.Url -like '*arm*'){
                                                $arch = "arm64"
                                            }elseif( $_.Url -like '*ia64*'){
                                                $arch = "ia64"
                                            }elseif( $_.Url -like '*x86*'){
                                                $arch = "x86"
                                            }elseif( $_.Url -like '*x64*'){
                                                $arch = "x64"
                                            }
                                            $this.awaitThreads()
                                            $filename = (split-path $_.Url -leaf)
                                            $this.logger("INFO", "Checking $filename" )
                                            Start-Job -ScriptBlock $downloadScriptBlock -ArgumentList ($_.Url , "$($this.execPath)\storage\$product\$arch\$filename", $this.reDownload)
                                        }
                                    }
                                }
                            }

                            get-job | ? { $_.State -eq 'Completed' } | remove-job
                            [GC]::Collect()
                        }
                    }

                    $this.awaitThreads(0)

                    $deleteNodes = $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[./dns:SupersededBy][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '$($prodId)']]",$ns)
                    if($deleteNodes -ne $null -and $deleteNodes.count -gt 1){
                        $deleteNodes | % {

                            $xml.selectnodes("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='$($_.RevisionId)'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id",$ns) | % {
                                $fileId = $($_.'#text')
                                $xml.selectnodes("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '$fileId']",$ns) | % {
                                    $fileName = $($_.Url.substring($_.url.lastindexof('/')+1))
                                    $this.logger("WARNING", "Superseding: $filename")
                                    get-childItem -Recurse -Path "./storage/$product/" -filter "*$($filename)*" | remove-item
                                }
                            }
                        }
                    }
                }
            }

            $this.awaitThreads(0)

            sleep -seconds 10
            get-job | ? { $_.State -eq 'Completed' } | remove-job
            [GC]::Collect()
            $this.logger("INFO", "Finished Downloading Windows Updates")
        }

    }
}
Process{
    $wsus = [wsusOfflineUpdater]::new($reDownload, $threads, $refreshDays)
}
End{

}
