#!/usr/bin/python3
import urllib.request
import subprocess
import libxml2
import os
import pprint
import time
import hashlib
import time
import stat
import logging
from urllib.parse import urlparse

class WindowsUpdater:
    BUF_SIZE = 65536
    
    logger = None
    
    products  = {
        # "msse"    : ["6cf036b9-b546-4694-885a-938b93216b66"],
        # "ofc03"   : ["1403f223-a63f-f572-82ba-c92391218055", "477b856e-65c4-4473-b621-a8b230bb70d9"],
        # "ofc07"   : ["041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591", "477b856e-65c4-4473-b621-a8b230bb70d9"],
        # "ofc10"   : ["84f5f325-30d7-41c4-81d1-87a0e6535b66", "477b856e-65c4-4473-b621-a8b230bb70d9"],
        # "ofc13"   : ["704a0a4a-518f-4d69-9e03-10ba44198bd5", "477b856e-65c4-4473-b621-a8b230bb70d9"],
        # "ofcLive" : ["03c7c488-f8ed-496c-b6e0-be608abb8a79", "477b856e-65c4-4473-b621-a8b230bb70d9"],
        # "wd"      : ["8c3fcc84-7410-4a95-8b89-a166a0190486"],        
        # "win50"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "3b4b8621-726e-43a6-b43b-37d07ec7019f"],
        # "win51"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","558f4bc3-4827-49e1-accf-ea79fd72d4c9", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94"],
        # "win52"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e", "dbf57a08-0d5a-46ff-b30c-7715eb9498e9", "032e3af5-1ac5-4205-9ae5-461b4e8cd26d", "a4bedb1d-a809-4f63-9b49-3fe31967b6d0", "4cb6ebd5-e38a-4826-9f76-1416a6f563b0", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "68623613-134c-4b18-bcec-7497ac1bfcb0"],
        # "win60"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","26997d30-08ce-4f25-b2de-699c36a8033a", "ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf", "575d68e2-7c94-48f9-a04f-4b68555d972d", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "6966a762-0c7c-4261-bd07-fb12b4673347", "e9b56b9a-0ca9-4b3e-91d4-bdcf1ac7d94d", "41dce4a6-71dd-4a02-bb36-76984107376d", "ec9aaca2-f868-4f06-b201-fb8eefd84cef", "68623613-134c-4b18-bcec-7497ac1bfcb0"],
        "win10"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","a3c2375d-0c8a-42f9-bce0-28333e198407", "d2085b71-5f1f-43a9-880d-ed159016d5c6", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94"],
        "win61"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","bfe5b177-a086-47a0-b102-097e4fa1f807", "f4b9c883-f4db-4fb5-b204-3343c11fa021", "fdfe8200-9d98-44ba-a12a-772282bf60ef", "1556fc1d-f20e-4790-848e-90b7cdbedfda", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94"],
        "win62"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","2ee2ad83-828c-4405-9479-544d767993fc", "a105a108-7c9b-4518-bbbe-73f0fe30012b", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4", "e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "0a07aea1-9d09-4c1e-8dc7-7469228d8195"],
        "win63"   : ["6964aab4-c5b5-43bd-a17d-ffb4346a8e1d","405706ed-f1d7-47ea-91e1-eb8860039715", "18e5ea77-e3d1-43b6-a0a8-fa3dbcd42e93", "6407468e-edc7-4ecd-8c32-521f64cee65e", "d31bd4c3-d872-41c9-a2e7-231f372588cb", "83aed513-c42d-4f94-b4dc-f2670973902d", "e6cf1350-c01b-414d-a61f-263d14d133b4","e0789628-ce08-4437-be74-2495b842f43b", "0fa1201d-4330-4fa8-8ae9-b877473b6441", "68c5b0a3-d1a6-4553-ae49-01d3a7827828", "28bc880e-0592-4cbf-8f95-c79b17911d5f", "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83", "56309036-4c77-4dd9-951a-99ee9c246a94", "2c62603e-7a60-4832-9a14-cfdfd2d71b9a"],
    }
        
    bad_languages = [
        "-af-za_", "-al_", "-am-et_", "-am_", "-arab-iq_", "-arab-pk_", "-as-in_", "-ba_", "-bd_", 
        "-be-by_", "-beta_", "-bg_", "-bgr_", "-bn-bd_", "-bn-in_", "-bn_", "-ca_", "-cat_", 
        "-chs_", "-cht_", "-cs-cz_", "-cs_", "-csy_", "-cym_", "-cyrl-ba_", "-cyrl-tj_", 
        "-da-dk_", "-da_", "-dan_", "-de-de_", "-de_", "-deu_", "-el-gr_", "-el_", "-ell_", 
        "-es-es_", "-es_", "-esn_", "-et_", "-eti_", "-eu_", "-euq_", "-fa-ir_", "-fi-fi_", 
        "-fi_", "-fil-ph_", "-fin_", "-fr-fr_", "-fr_", "-fra_", "-gd-gb_", "-ge_", "-ger_", 
        "-gl_", "-glc_", "-gu-in_", "-hbr_", "-he-il_", "-he_", "-heb_", "-hi_", "-hin_", "-hk_", 
        "-hr_", "-hrv_", "-hu-hu_", "-hun_", "-hy-am_", "-id_", "-ig-ng_", "-in_", "-ind_", 
        "-ir_", "-ire_", "-is-is_", "-is_", "-isl_", "-it-it_", "-it_", "-ita_", "-ja-jp_", 
        "-jpn_", "-ka-ge_", "-ke_", "-kg_", "-kh_", "-km-kh_", "-kn-in_", "-ko-kr_", "-kok-in_", 
        "-kor_", "-ky-kg_", "-latn-ng_", "-latn-uz_", "-lb-lu_", "-lbx_", "-lk_", "-lt_", "-lth_", 
        "-lu_", "-lv_", "-lvi_", "-mi-nz_", "-ml-in_", "-mlt_", "-mn-mn_", "-mn_", "-mr-in_", 
        "-ms-bn_", "-msl_", "-mt-mt_", "-mt_", "-nb-no_", "-nb_", "-ne-np_", "-ng_", "-nl-nl_", 
        "-nl_", "-nld_", "-nn-no_", "-nn_", "-no_", "-non_", "-nor_", "-np_", "-nso-za_", "-nz_", 
        "-or-in_", "-pa-in_", "-pe_", "-ph_", "-pk_", "-pl-pl_", "-pl_", "-plk_", "-pt-br_", 
        "-pt-pt_", "-ptb_", "-ptg_", "-qut-gt_", "-quz-pe_", "-ro-ro_", "-ro_", "-rom_", 
        "-ru-ru_", "-ru_", "-rus_", "-rw-rw_", "-si-lk_", "-sk-sk_", "-sk_", "-sky_", "-sl-si_", 
        "-sl_", "-slv_", "-sq-al_", "-srl_", "-sv-se_", "-sv_", "-sve_", "-sw-ke_", "-ta-in_", 
        "-te-in_", "-tha_", "-ti-et_", "-tk-tm_", "-tm_", "-tn-za_", "-tr-tr_", "-tr_", "-trk_", 
        "-tt-ru_", "-ug-cn_", "-uk-ua_", "-uk_", "-ukr_", "-ur-pk_", "-uz_", "-vit_", "-wo-sn_", 
        "-xh-za_", "-yo-ng_", "-za_", "-zh-cn_", "-zh-hk_", "-zh-tw_", "-zhh_", "-zu-za_", 
        "-af-za_", "-ar-sa_", "-ar_", "-ara_", "-az-latn-", "-bg-bg_", "-bs-latn", "-ca-es", 
        "-cy-gb", "-et-ee", "-eu-es", "-ga-ie", "-gl-es", "-hi-in", "-hr-hr", "-id-id", "-kk-kz", 
        "-lt-lt", "-lv-lv", "-mk-mk", "-ms-my", "-prs-af", "-sr-cyrl", "-sr-latn", "-th-th", 
        "-vi-vn"
    ]
        
        
    def __init__(self):
        self.logger = logging.getLogger('microsoft_updates')
        self.logger.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        ch.setFormatter(formatter)
        
        fh = logging.FileHandler("microsoft_updates.log")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        
        self.logger.addHandler(ch)
        self.logger.addHandler(fh)
        
        self.logger.debug('Starting app')
        
    def file_age_in_seconds(self, pathname):
        if os.path.exists(pathname):
            return time.time() - os.stat(pathname)[stat.ST_MTIME]
        else:
            return 0
    
    def download_file(self, product, arch, url):
        filename = os.path.basename(url)
        if not os.path.exists('./storage/{}/{}/{}'.format(product, arch, filename) ):
            self.logger.info("\tDownloading {}".format( filename ) )
            urllib.request.urlretrieve(url, "./storage/{}/{}/{}".format(product, arch, filename) )
        else:
            sha1 = hashlib.sha1()
            data = None
            with open('./storage/{}/{}/{}'.format(product, arch, filename) , 'rb') as f:
                while True:
                    data = f.read(self.BUF_SIZE)
                    if not data:
                        break
                    sha1.update(data)
            if sha1.hexdigest() in filename:
                self.logger.info("\tUsing existing {} with hash {}".format( filename, sha1.hexdigest() ))
            else:
                self.logger.warning("\tRe-Downloading {} due to hash mismatch".format( filename ))
                os.remove('./storage/{}/{}/{}'.format(product, arch, filename) )
                urllib.request.urlretrieve(url, "./storage/{}/{}/{}".format(product, arch, filename) )
                
                
    def download_wsus_agent(self):
        if not os.path.exists("./storage/wua/"):
            os.mkdir('./storage/wua/')
        if not os.path.exists("./storage/wua/x64/"):
            os.mkdir('./storage/wua/x64/')
        if not os.path.exists("./storage/wua/x86/"):
            os.mkdir('./storage/wua/x86/')
        if not os.path.exists("./storage/wua/x86_x64/"):
            os.mkdir('./storage/wua/x86_x64/')
            
        links = [
            ('http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab', './storage/wua/wsusscn2.cab'),
            ('http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/windowsupdateagent30-x86.exe', './storage/wua/x86/windowsupdateagent30-x86.exe'),
            ('http://download.windowsupdate.com/windowsupdate/redist/standalone/7.4.7600.226/windowsupdateagent30-x64.exe', './storage/wua/x64/windowsupdateagent30-x64.exe')
        ]
        
        for link in links:
            if 0 < self.file_age_in_seconds( link[1] ) < 86400:
                self.logger.info('Using existing {}'.format( os.path.basename(link[1])))
            else:
                if os.path.exists( link[1]):
                    self.logger.warning('Removing old {}'.format( link[1]) )
                    os.remove(link[1])
                self.logger.info('Downloading {}'.format(os.path.basename(link[1])))
                urllib.request.urlretrieve(link[0], link[1])

        if os.path.exists("./storage/wua/package.cab"):
            self.logger.warning('Removing old package.cab')
            os.remove("./storage/wua/package.cab")

        self.logger.info('Extracting package.cab from wsusscn2.cab')
        processes = subprocess.check_output(
            [ './bin/7z', 'e', './storage/wua/wsusscn2.cab', 'package.cab', '-o./storage/wua/' ],
            stderr=None
        ).decode('ascii').split("\n")

        if os.path.exists("./storage/wua/package.xml"):
            self.logger.warning('Removing old package.xml')
            os.remove("./storage/wua/package.xml")

        self.logger.info('Extracting pacakge.xml from package.cab')
        processes = subprocess.check_output(
            [ './bin/7z', 'e', './storage/wua/package.cab', 'package.xml', '-o./storage/wua/' ],
            stderr=None
        ).decode('ascii').split("\n")

        self.logger.info('Generating pretty xml')
        command = "cat ./storage/wua/package.xml | xmllint --format - > ./storage/wua/formated_package.xml"
        output = subprocess.check_output(["bash", "-c", command])

    def download_windows_defender(self):
        if not os.path.exists("./storage/wd/"):
            os.mkdir('./storage/wd/')
        if not os.path.exists("./storage/wd/x64/"):
            os.mkdir('./storage/wd/x64/')
        if not os.path.exists("./storage/wd/x86/"):
            os.mkdir('./storage/wd/x86/')
            
        links = [
            ('https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64', './storage/wd/x64/mpam-feX64.exe'),
            ('https://go.microsoft.com/fwlink/?LinkID=121721&arch=x86', './storage/wd/x86/mpam-feX86.exe'),
        ]
        
        for link in links:
            if 0 < self.file_age_in_seconds( link[1]) < 86400:
                self.logger.info('Using existing {}'.format( os.path.basename(link[1])))
            else:
                if os.path.exists( link[1]):
                    self.logger.warning('Removing old {}'.format( link[1]) )
                    os.remove(link[1])
                self.logger.info('Downloading {}'.format(os.path.basename(link[1])))
                urllib.request.urlretrieve(link[0], link[1])

    def download_cpp_dotnet(self):
        if not os.path.exists("./storage/cpp_dn/"):
            os.mkdir('./storage/cpp_dn/')
        if not os.path.exists("./storage/cpp_dn/x64/"):
            os.mkdir('./storage/cpp_dn/x64/')
        if not os.path.exists("./storage/cpp_dn/x86/"):
            os.mkdir('./storage/cpp_dn/x86/')
        if not os.path.exists("./storage/cpp_dn/x86_x64/"):
            os.mkdir('./storage/cpp_dn/x86_x64/')
            
        links = [
            ('http://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.EXE', './storage/cpp_dn/x86/vcredist2005_x86.exe'),
            ('http://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe', './storage/cpp_dn/x86/vcredist2008_x86.exe'),
            ('http://download.microsoft.com/download/E/E/0/EE05C9EF-A661-4D9E-BCE2-6961ECDF087F/vcredist_x86.exe', './storage/cpp_dn/x86/vcredist2010_x86.exe'),
            ('http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe', './storage/cpp_dn/x86/vcredist2012_x86.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/10912113/5da66ddebb0ad32ebd4b922fd82e8e25/vcredist_x86.exe', './storage/cpp_dn/x86/vcredist2013_x86.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/9307e627-aaac-42cb-a32a-a39e166ee8cb/E59AE3E886BD4571A811FE31A47959AE5C40D87C583F786816C60440252CD7EC/VC_redist.x86.exe', './storage/cpp_dn/x86/vcredist2019_x86.exe'),

            ('http://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE', './storage/cpp_dn/x64/vcredist2005_x64.exe'),
            ('http://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe', './storage/cpp_dn/x64/vcredist2008_x64.exe'),
            ('http://download.microsoft.com/download/E/E/0/EE05C9EF-A661-4D9E-BCE2-6961ECDF087F/vcredist_x64.exe', './storage/cpp_dn/x64/vcredist2010_x64.exe'),
            ('http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe', './storage/cpp_dn/x64/vcredist2012_x64.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/10912041/cee5d6bca2ddbcd039da727bf4acb48a/vcredist_x64.exe', './storage/cpp_dn/x64/vcredist2013_x64.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/3b070396-b7fb-4eee-aa8b-102a23c3e4f4/40EA2955391C9EAE3E35619C4C24B5AAF3D17AEAA6D09424EE9672AA9372AEED/VC_redist.x64.exe', './storage/cpp_dn/x64/vcredist2019_x64.exe'),
            
            ('http://download.windowsupdate.com/c/msdownload/update/software/updt/2017/10/ndp46-kb4041778-x86_7eb6b6850959824025c3d3e217391bcb512b30c7.exe','./storage/cpp_dn/x86/ndp46-kb4041778-x86.exe'),
            ('http://download.windowsupdate.com/c/msdownload/update/software/updt/2017/10/ndp46-kb4041778-x64_1e36c362f9f6c4dfdef1a3c0abe4fb560ec7f25c.exe','./storage/cpp_dn/x64/ndp46-kb4041778-x64.exe'),
            ('http://download.microsoft.com/download/2/0/e/20e90413-712f-438c-988e-fdaa79a8ac3d/dotnetfx35.exe','./storage/cpp_dn/x86_x64/dotnetfx35.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/014120d7-d689-4305-befd-3cb711108212/0fd66638cde16859462a6243a4629a50/ndp48-x86-x64-allos-enu.exe','./storage/cpp_dn/x64/ndp48-x86-x64-allos-enu.exe'),
            ('http://download.visualstudio.microsoft.com/download/pr/7afca223-55d2-470a-8edc-6a1739ae3252/65592cace88fcfb3e14a5c4a54833794/ndp48-x86-x64-allos-deu.exe','./storage/cpp_dn/x86/ndp48-x86-x86-allos-enu.exe')
        ]
        
        for link in links:
            if 0 < self.file_age_in_seconds( link[1]) < 86400:
                self.logger.info('Using existing {}'.format( os.path.basename(link[1])))
            else:
                if os.path.exists( link[1]):
                    self.logger.warning('Removing old {}'.format( link[1]) )
                    os.remove(link[1])
                self.logger.info('Downloading {}'.format(os.path.basename(link[1])))
                urllib.request.urlretrieve(link[0], link[1])
    
    def download_updates(self):
        package = libxml2.parseFile("./storage/wua/package.xml")
        
        package_context = package.xpathNewContext()
        package_context.xpathRegisterNs("dns", "http://schemas.microsoft.com/msus/2004/02/OfflineSync")
        
        for product  in self.products.keys():
            if not os.path.exists("./storage/{}/".format(product)):
                os.mkdir('./storage/{}/'.format(product) )
            
            self.logger.info( "Processing Updates for {}".format( product ) )
            
            for arch in ["x86","x64","x86_x64"]:
                if not os.path.exists("./storage/{}/{}/".format(product, arch)):
                    os.mkdir('./storage/{}/{}/'.format(product, arch) )
            
            for product_id in self.products[product]:
                self.logger.info("Sub Product ID: {}".format(product_id) )
                product_nodes = package_context.xpathEval("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '{}']]".format(product_id))
                
                
                pni = 0
                pnt = len(product_nodes)
                for product_node in product_nodes:
                    pni += 1
                    
                    update_id = (product_node.xpathEval('@UpdateId')[0].content)
                    revision_number = (product_node.xpathEval('@RevisionNumber')[0].content)
                    revision_id = (product_node.xpathEval('@RevisionId')[0].content)
    
                    self.logger.info("{} / {}: Update:{}, Revision: {}, RevisionId: {}".format(pni, pnt, update_id, revision_number, revision_id) )
    
                    update_nodes = package_context.xpathEval(
                        "//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(./dns:SupersededBy)][not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='{}'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id".format(revision_id)
                    )

                    for update_node in update_nodes:
                        file_id = update_node.content
                        self.logger.info("\tFile ID:{}".format(file_id) )
                        file_nodes = package_context.xpathEval("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '{}']/@Url".format(file_id))
                        for file_node in file_nodes:
                            self.logger.info("\tChecking: {}".format(file_node.content) )
                            if 'mui' not in file_node.content and not any(x in file_node.content for x in self.bad_languages):
                                if 'x86' in file_node.content and 'x64' not in file_node.content:
                                    self.download_file(product, 'x86', file_node.content)
                                elif 'x64' in file_node.content and 'x86' not in file_node.content:
                                    self.download_file(product, 'x64', file_node.content)
                                else:
                                    self.download_file(product, 'x86_x64', file_node.content)
                
                delete_product_nodes = package_context.xpathEval("//dns:OfflineSyncPackage/dns:Updates/dns:Update[./dns:SupersededBy][./@DefaultLanguage='en' or not(./@DefaultLanguage)][./@IsBundle='true'][./dns:Categories/dns:Category[./@Type = 'Product' and ./@Id = '{}']]".format(product_id))
                index = 0
                total_nodes = len( delete_product_nodes )
                for delete_product_node in delete_product_nodes:
                    index += 1
                    delete_revision_id = (delete_product_node.xpathEval('@RevisionId')[0].content)
                    
                    delete_nodes =  package_context.xpathEval("//dns:OfflineSyncPackage/dns:Updates/dns:Update[not(@isBundle)][./dns:BundledBy/dns:Revision/@Id='{}'][not(./dns:Languages) or ./dns:Languages/dns:Language/@Name='en']/dns:PayloadFiles/dns:File/@Id".format(delete_revision_id))
                    for delete_node in delete_nodes:
                        delete_file_id = delete_node.content
                        delete_file_nodes = package_context.xpathEval("//dns:OfflineSyncPackage/dns:FileLocations/dns:FileLocation[@Id = '{}']/@Url".format(delete_file_id))
                        for delete_file_node in delete_file_nodes:
                            self.logger.info("{} / {}: Superseding {}".format( index, total_nodes, os.path.basename( delete_file_node.content ) ) )
                            
                            if os.path.exists("./storage/{}/x86/{}".format(product, os.path.basename( delete_file_node.content ))): 
                                os.remove("./storage/{}/x86/{}".format(product, os.path.basename( delete_file_node.content )))
                            elif os.path.exists("./storage/{}/x64/{}".format(product, os.path.basename( delete_file_node.content ))): 
                                os.remove("./storage/{}/x64/{}".format(product, os.path.basename( delete_file_node.content )))
                            elif os.path.exists("./storage/{}/x86_64/{}".format(product, os.path.basename( delete_file_node.content ))): 
                                os.remove("./storage/{}/x86_x64/{}".format(product, os.path.basename( delete_file_node.content )))
                            
    
if __name__ == '__main__':
    wu = WindowsUpdater()
    
    wu.download_wsus_agent()
    wu.download_windows_defender()
    wu.download_cpp_dotnet()
    wu.download_updates()
