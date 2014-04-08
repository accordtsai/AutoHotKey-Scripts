; IMPORTANT INFO ABOUT GETTING STARTED: Lines that start with a
; semicolon, such as this one, are comments.  They are not executed.

; This script has a special filename and path because it is automatically
; launched when you run the program directly.  Also, any text file whose
; name ends in .ahk is associated with the program, which means that it
; can be launched simply by double-clicking it.  You can have as many .ahk
; files as you want, located in any folder.  You can also run more than
; one .ahk file simultaneously and each will get its own tray icon.

; SAMPLE HOTKEYS: Below are two sample hotkeys.  The first is Win+Z and it
; launches a web site in the default browser.  The second is Control+Alt+N
; and it launches a new Notepad window (or activates an existing one).  To
; try out these hotkeys, run AutoHotkey again, which will load this file.

; #z::Run www.autohotkey.com

^!n::
IfWinExist Untitled - Notepad
	WinActivate
else
	Run Notepad
return

; Note: From now on whenever you run AutoHotkey directly, this script
; will be loaded.  So feel free to customize it to suit your needs.

; Please read the QUICK-START TUTORIAL near the top of the help file.
; It explains how to perform common automation tasks such as sending
; keystrokes and mouse clicks.  It also explains more about hotkeys.

; True passwords have been replaced by the string 'ppppwwww'. Replace this text
; with your true password to make this script useful.

;=================================
; PACS, RT Chart, Notepad, mspaint
;=================================
#x::
IfWinActive INFINITT PACS
	{
	WinMinimize
	return
	}
IfWinExist INFINITT PACS
	WinActivate
else
{
	InputBox, MRD, 輸入病歷號, 病歷號:
	Run iexplore.exe http://172.17.2.2/pkg_pacs/external_interface.aspx?&LID=051007&LPW=ppppwwww&PID=%MRD%
}
return

; RT Chart
#z::
Run iexplore.exe C:\Users\NCKU\RTchartAutoLogin.html
return

; Notepad
#n::
IfWinExist 未命名 - 記事本
	WinActivate
else
	Run Notepad
return

#c:: Run Calc

; EMR
; Temporary Ctrl+C for copying MRD
#e::

;'IfWinExist Web EMR Login Page
;	WinClose
;IfWinExist Web EMR
;	{
;	WinActivate
;	return
;	}
;else
;	{
;	Run http://hisweb.hosp.ncku/Emrquery/JqueryUI.aspx
;	InputBox, MRD, 輸入病歷號, 病歷號:
;	Run iexplore.exe http://hisweb.hosp.ncku/Emrquery/autologin.aspx?chartno=%MRD;

	Run iexplore.exe http://hisweb.hosp.ncku/Emrquery/autologin.aspx?chartno=%clipboard%

;IfWinExist Web EMR Login Page
;	{
;	Sleep 5000
;	Send {Tab}{Tab}{Tab}{Tab}{Tab}051007{Tab}n@ppppwwww{Enter}
;	}
;	}
return

;; 開 SOAP --> 複製病歷號開 EMR 跟 PACS --> 點 Ditto
;; 要先把左邊病患清單固定住，寬度調到剛好

^s::
IfWinActive 成大醫院HIS系統
{
	Click 500,73   		;開立醫囑
	Click 660,100,2		;病歷號
	Send ^c
        Click 550,180		;SOAP
        Click 600,675		;Ditto
	Run iexplore.exe http://hisweb.hosp.ncku/Emrquery/autologin.aspx?chartno=%clipboard%
;	Run iexplore.exe http://172.17.2.2/pkg_pacs/external_interface.aspx?&LID=051007&LPW=ppppwwww&PID=%clipboard%
;	Sleep 5000
;	WinWait, Web EMR,%clipboard%,5
;	if ErrorLevel
;		return
;	else
;		Click 2700,110	;一年
	
}
return

#p::Run mspaint

;===============
; PCS 開會診回覆
;===============
; 先打開，如果有更新完之後，Cursor 到帳號的地方按 Ctrl-Alt-C
; PCS 要全螢幕
^!c::
IfWinActive 成醫醫療資訊系統
{
	Send 051007{Tab}n@ppppwwww{Tab}9200{Enter}
	Sleep 7000
	WinActivate, 成大醫院PCS系統
	Click 70,40	; 系統功能
	Click 100,260	; 住院醫囑
	Click 300,260	; 開立醫囑
	Click 500,290	; 會診回覆
	Sleep 2000
;	WinActivate, 成大醫院PCS系統
;	Click 425,100	; 待回覆
	Click 400,100	; 待回覆 (新位置)
	Click 535,150	; 類別
	Click 400,150	; 類別 (新位置)
	Send {Down}{Down}{Down}	;依被會科別查詢
	Click 730,150	; 隔壁
	Send 92{Enter}
	Click 900,150
}
	
;====
;HIS
;====
^!h::
IfWinActive 成醫醫療資訊系統
{
	Send 051007{Tab}n@ppppwwww{Enter}
}

;========
;懶惰鬼
;========
;; Breast & GYN abbreviations
::mam::Mammography
::bson::Breast sonography
::paps::
Send Pap smear{Space}
Gosub, DateHelper
Send Negative
return
::gyns::
Send GYN sonography
Gosub, DateHelper
Send Endometrium ??? cm
return
::abec::Abdominal echo
::mamfu::
InputBox, date_y, Enter Year, 201 (3 or 4):
InputBox, date_m, Enter Date, Month:
InputBox, date_d, Enter Date, Day:
Send Mammography (201%date_y%/%date_m%/%date_d%):{space}{enter}
Send Breast sonography (201%date_y%/%date_m%/%date_d%):{space}{enter}
Send CXR (201%date_y%/%date_m%/%date_d%):{space}{enter}
Send Abdominal echo (201%date_y%/%date_m%/%date_d%):{space}
return
::br0::(BIRADs 0)
::br1::(BIRADs 1)
::br2::(BIRADs 2)
::br3::(BIRADs 3)
::br4::(BIRADs 4)
::br5::(BIRADs 5)
::br6::(BIRADs 6)
:O:schd::scheduled on 2014/ by {left 4}
::arrd::arranged by
:O:patho::Pathology (


^d::
Send (2013//):{Space}{left 4}
;else
;Send ^d
return

^q::Send (2014//):{Space}{left 4}


;; CT & MR abbreviations
::mri::MRI
::cxr::CXR
::bct::Brain CT
::hnct::Head and neck CT
::cct::Chest CT
::abct::Abdominal CT
::pvct::Pelvic CT
::wbct::Whole body CT
::wct::Whole body CT
::spmr::Spine MRI
::brmr::Brain MRI
::hnmr::Head and neck MRI
::pvmr::Pelvic MRI
::abmr::Abdominal MRI
::petct::PET-CT
::bsc::Bone scan

::cfs::Colonoscope
::pes::PES
::egd::EGD
::npy::Nasopharyngoscopy
::fby::Fibroscopy

; Tumor markers
::cea::CEA:
::c9::CA-199:
::c3::CA-153:
::c5::CA-125:
::afp::AFP:
;::psa::PSA:


;; Enter date
DateHelper:
InputBox, date_y, Enter Year, 201 (3 or 4):
InputBox, date_m, Enter Date, Month:
InputBox, date_d, Enter Date, Day:
Send (201%date_y%/%date_m%/%date_d%):{Space}
return

;; Other common expression when writing chart
:O:cprt::C.C.: For post-R/T follow-up{enter}Completed R/T on 2014/.{left 1}
:O:strt::C.C.: For status check during R/T treatment{enter}Started R/T on 2014/.{left 1}
::cur::Current R/T given dose:
::fx::fractions
::cgy::cGy
::gy::Gy
::30gy::3000 cGy / 10 fractions / 2 weeks
::d/::cGy / fractions /  days{left 18}
::w/::cGy / fractions /  weeks{left 19}
::ms::months
::ys::years
::wks::weeks
::ds::days
:c:lt::left
:c:rt::right
:c:Lt::Left
:c:Rt::Right
::wbrt::WBRT
::1w::Go on R/T,{enter}RTC 1 week
::rtc::RTC
::r1y::RTC 1 year
::r6m::RTC 6 months
::r3m::RTC 3 months
::r1m::RTC 1 month
::r4w::RTC 4 weeks
::a/b::α/β
::hbhct::Hb/Hct/WBC/Plt:
::pend::[ pending ]
::adjrt::adjuvant radiotherapy
::palrt::palliative radiotherapy
::brca::Breast cancer
::cxca::Cervical cancer
::supc::supportive care
::kfu::Keep follow-up
::cdsim::Cone down simulation
:?:/*-::(-)
:?:56+::(`{+})
::rtp::R/T plan:
::x6c::x6 cycles
::x4c::x4 cycles

::rae::
(
Radiotherapy Adverse Effect
Fatigue           Gr 0, Radiation dermatitis Gr 0
Diarrhea          Gr 0, Proctitis            Gr 0
Enteritis/Colitis Gr 0, Dysphagia            Gr 0
Cough             Gr 0, Dyspnea              Gr 0
Mucositis         Gr 0, Xerostomia           Gr 0
Dysuria           Gr 0, Incontinence         Gr 0
)

;; Medication
::1q7::1# qd x 7 days
::1b7::1# bid x 7 days
::1t7::1# tid x 7 days
::2q7::2# qd x 7 days
::2b7::2# bid x 7 days
::2t7::2# tid x 7 days
::1q14::1# qd x 14 days
::1b14::1# bid x 14 days
::1t14::1# tid x 14 days
::2q14::2# qd x 14 days
::2b14::2# bid x 14 days
::2t14::2# tid x 14 days
::1q21::1# qd x 21 days
::1b21::1# bid x 21 days
::1t21::1# tid x 21 days
::2q21::2# qd x 21 days
::2b21::2# bid x 21 days
::2t21::2# tid x 21 days

::acet::Acetaminophen
::cata::Cataflam
::flum::Acetylcysteine
::pred::Predonine
::dori::Dorison
::rowa::Rowapraxin

;; Doctor names
::tmh::蔡牧宏
::drhhw::Dr. 陳海雯
::drlfj::Dr. 林逢嘉
::drxwt::Dr. 薛尉廷
::drlyx::Dr. 賴俞璇
::drwyh::Dr. 吳沅樺

::drswj::Dr. 蘇五洲
::drswb::Dr. 蘇文彬
::drwsy::Dr. 吳尚殷
::drcyp::Dr. 陳雅萍
::dryym::Dr. 葉裕民
::dryjr::Dr. 顏家瑞
::drzwb::Dr. 鍾為邦

::drlgd::Dr. 李國鼎
::drgyl::Dr. 郭耀隆
::drlbw::Dr. 林博文
::drccw::Dr. 張財旺
::drlzc::Dr. 李政昌
::drxhp::Dr. 徐慧萍
::drlyj::Dr. 李宜堅
::drwlc::Dr. 王亮超

::drzzn::Dr. 鄭兆能

::dryyt::Dr. 顏亦廷
::drlww::Dr. 賴吾為

::drtst::Dr. 蔡森田
::drxzr::Dr. 蕭振仁
::drojy::Dr. 歐俊巖
::drlhy::Dr. 羅宏一
::drlxz::Dr. 廖筱蓁

::drzym::Dr. 鄭雅敏
::drzzy::Dr. 周振陽
::drhyf::Dr. 黃于芳

::drojh::Dr. 歐建慧
::drtzx::Dr. 蔡宗欣

;; New patient
::newpatient::
InputBox, refdoc, Referred from, Referred from:
InputBox, reffor, for, for:
InputBox, age, Input Age, Age:
InputBox, sex, Input Sex (Please enter 'fe' or blank), Sex:
if ErrorLevel  ; The user pressed Cancel.
    return
Send Chief complaint: Referred from %refdoc%
Send for %reffor%{Enter}{Enter}Review of history:{Enter}1) %age%-year-old %sex%male. PHx: DM(-) HTN(-). OP Hx: Nil.{Enter}2) 20{Enter}Referred to our department for %reffor%{Enter}
Click 500,500   		;Objective
Send BW: Kg{Enter}PS = 0
return

;; Consultation
::newcons::
(
Dear Dr. 黃 & Dr. 蘇:

We have reviewed history and images of this patient.

Review of history:


Suggestion:
Adjuvant/Palliative radiotherapy is indicated. We have visited the patient and family, and explained risks and benefits of radiotherapy. Pre-R/T preparations have been arranged on 2014/. Please send the patient and chart to Radiation Oncology department (住院大樓 B2) on time. Radiotherapy will start as soon as treatment planning is complete.

Thank you for your consultation!
R 蔡牧宏 (7501) / VS 薛尉廷 (7534)
)

;; Physical exam
::normalpe::
(
HEENT: no gross deformities, conjuctiva: not anemic, sclera: not icteric
Neck: supple, no jugular vein engorgement, no lymphadenopathy
Chest: breathing sounds: clear, symmetrical expansion
Abdomen: soft and flat, liver/spleen: impalpable
Extremities: freely movable, no pitting edema, petechia(-), ecchymosis(-)
)

::breastpe::
(
HEENT: no gross deformities, conjuctiva: not anemic, sclera: not icteric
Neck: supple, no jugular vein engorgement, no lymphadenopathy
Chest: breathing sounds: clear, symmetrical expansion; 
Breast: s/p OP, as picture
Abdomen: soft and flat
Extremities: freely movable, petechia(-), ecchymosis(-)
)

;; ER/PR status
::erpr::
InputBox, er, ER, ER[+ or -]:
InputBox, ern, ER, ER[`%]:
InputBox, pr, PR, PR[+ or -]:
InputBox, prn, PR, PR[`%]:
InputBox, neu, Her-2/Neu, Her-2/Neu:
InputBox, ki, Ki-67, Ki-67[`%]
Send ER(%er%)(%ern%`%) PR(%pr%)(%prn%`%) Her-2/Neu(%neu%) Ki-67: %ki%`%
return

;=============
;Brachytherapy
;=============

::compcc::
(
1st intracavitary brachytherapy was performed smoothly.

Uterine sounding: xxx cm.
Intracavitary brachytherapy was performed with full ovoids.
Dose: xxx cGy.

Rx: 
N/S 1 BT for foley irrigation
Lomexin 1# Supp QD x 3d
Vaginal irrigation
On Foley catheter and remove after treatment

Plan:
Continue EBRT
next ICRT on 2014/4/10
)

::simpcc::
(
1st vaginal mould intracavitary brachytherapy was performed smoothly. 

D = 3 cm, L = 3 cm. 

Rx:
300 cGy 
Next brachytherapy on 2014/4/10.
)

;========
;R/T plan
;========

::rtrbcs::
(
R/T plan:
Right whole breast: 5000 cGy/25 fractions/5 weeks
Right breast tumor bed boost: 6400 cGy/32 fractions/6-7 weeks
)

::rtlbcs::
(
R/T plan:
Left whole breast: 5000 cGy/25 fractions/5 weeks
Left breast tumor bed boost: 6400 cGy/32 fractions/6-7 weeks
)

::rtrmrm::
(
R/T plan:
Right supra- & infra-clavicular LNs & right anterior chest wall: 
5000 cGy / 25 fractions / 5 weeks
)

::rtrmrme::
(
R/T plan:
Right supra- & infra-clavicular LNs & right anterior chest wall: 5000 cGy / 25 fractions / 5 weeks
Right anterior chest wall scar boost: 6000 cGy / 30 fractions / 6 weeks
)

::rtlmrm::
(
R/T plan:
Left supra- & infra-clavicular LNs & left anterior chest wall: 
5000 cGy / 25 fractions / 5 weeks
)

::rtlmrme::
(
R/T plan:
Left supra- & infra-clavicular LNs & left anterior chest wall: 5000 cGy / 25 fractions / 5 weeks
Left anterior chest wall scar boost: 6000 cGy / 30 fractions / 6 weeks
)

::rtrecpost::
(
R/T plan:
Pelvic regional LNs and rectal tumor surgical bed: 
5040 cGy/ 28 fractions / 5-6 weeks
)

::rtcxpost::
(
R/T plan:
Pelvic regional LNs & vaginal stump: 4500 cGy/25 fractions/5 weeks
Vaginal intracavitary brachytherapy (submucosal 0.5 cm): 300 cGy x 5 fractions
)

::rtcxca::
(
R/T plan:
Whole pelvis: 4500 to 5040 cGy/25-28 fractions/5-6 weeks (adding central BSB after 3960 cGy)
High-dose-rate intracavitary brachytherapy (ICRT)(point A): 400-500 cGy x 6-8 fractions
)

::rtwbrt::
(
R/T plan:
Whole brain: 3000 cGy / 10 fractions / 2 weeks
)

;=================================================================
; cGy / fractions /  days
;
; Defaults to adding 5 fractions and 7 days; adjust settings here! 
;=================================================================

#w::

augfx=5
augdays=7

;augfx=4
;augdays=7

Send ^c
FoundPos := RegExMatch(clipboard, "([0-9]+) cGy", olddose)
;FoundPos := RegExMatch(clipboard, "([0-9]+) fractions", oldfx)
FoundPos := RegExMatch(clipboard, "([0-9]+) fx", oldfx)
FoundPos := RegExMatch(clipboard, "([0-9]+) days", olddays)
fxsize := olddose1 // oldfx1
newdose := olddose1 + fxsize * augfx
newfx := oldfx1 + augfx
newdays := olddays1 + augdays
Send %newdose% cGy / %newfx% fx / %newdays% days
return
