                    
# Win 1250 encoding

$Tab = [char]9
$i = 0

foreach($UserObj in (Get-ADuser -properties msExchRecipientTypeDetails,displayname,givenname,sn,title,department,mobile,mail,c,UserPrincipalName -Filter {(enabled -eq "True") -and ((msExchRecipientTypeDetails -eq "34359738368") -or (msExchRecipientTypeDetails -eq "2147483648"))} -SearchBase 'DC=int,DC=domain,DC=sk' | Sort-Object Displayname)){

$i = $i + 1
#$sUserName = $UserObj.sAMAccountName
$sFullName = $UserObj.displayname
$sMeno = $UserObj.givenName
$sPriezvisko = $UserObj.sn
#'sGrupy = UserObj.groups
#$sTitle = $UserObj.initials
$sPosition = $UserObj.title
$sDepartment = $UserObj.department
$sMobile = $UserObj.mobile
$sEmail = $UserObj.mail
$sLocation = $UserObj.c


$sWeb = "www.web.sk"
$logo_src = "TMR_Signature_"
$sCompany = "Firm, a.s."
$sStreetAddress = "Adress 72"
$sPostalCode = "03101"
$sStreet = "Lipláš"
$sState = "Slovakia"
$sPozdav = "--"


$sMobile = $sMobile -replace "[^0-9\+]" , "" # vyberie len čísla a znak "+"

switch ($sLocation) {
	"SK" {
            $sPozdav = "S pozdravom"; 
            if($sMobile.length -lt 11){ $sMobile = $sMobile -replace '^0', '+421' }
            if($sMobile.length -gt 12){ $($sMobile = $sMobile.SubString(0,4) + " " + $sMobile.SubString(4,3)  + " " + $sMobile.SubString(7,3) + " " + $sMobile.SubString(10,3)); }
            break}
	"CZ" {
            $sPozdav = "S pozdravem";
            if($sMobile.length -lt 11){ $sMobile = $sMobile -replace '^0', '+420' }
            if($sMobile.length -gt 12){ $($sMobile = $sMobile.SubString(0,4) + " " + $sMobile.SubString(4,3)  + " " + $sMobile.SubString(7,3) + " " + $sMobile.SubString(10,3)); }
            break}
	"AT" {
            $sPozdav = "Herzliche Grüße";
            if($sMobile.length -lt 10){ $sMobile = $sMobile -replace '^0', '+43' }
            #if($sMobile.length -eq 12){ $($sMobile = $sMobile.SubString(0,3) + " " + $sMobile.SubString(3,3)  + " " + $sMobile.SubString(6,3) + " " + $sMobile.SubString(9,3)); }
            if($sMobile.length -gt 11){ $($sMobile = $sMobile.SubString(0,3) + " " + $sMobile.SubString(3,3)  + " " + $sMobile.SubString(6,3) + " " + $sMobile.SubString(9,3) + " " + $sMobile.SubString(12)); }
            break}
	"PL" {
            $sPozdav = "Pozdrawiam";
            if($sMobile.length -lt 11){ $sMobile = $sMobile -replace '^0', '+48' }
            if($sMobile.length -gt 12){ $($sMobile = $sMobile.SubString(0,3) + " " + $sMobile.SubString(3,3)  + " " + $sMobile.SubString(6,3) + " " + $sMobile.SubString(9,3)); }
            break}
	"EN" {
            $sPozdav = "Kind Regards";
            break}

	default {$sPozdav = "--"; break}

}

foreach($GroupObj in (Get-ADPrincipalGroupMembership -Identity $UserObj | Where-Object {$_.name -like "TMR_Signature_*"} )){


switch ($GroupObj.name) {

	"TMR_Signature_CZCZYRK" {
		$logo_src = "TMR_Signature_CZCZYRK ";
    	$sWeb = "www.hotels.com";
      	$sCompany = " Bla bla S.A.";
		$sStreetAddress = "Narc 10";
		$sPostalCode = "43370";
		$sStreet = "Bla Bla";
		$sState = "Poland";
		break
	}

    "TMR_Signature_GRO" {
      	$logo_src = "TMR_Signature_GRO";
      	$sWeb = "www.golf.cz";
      	$sCompany = "TČR, a.s.";
		$sStreetAddress = "Oste 75";
		$sPostalCode = "73914";
		$sStreet = "Ostre";
		$sState = "Czech republic";
		break
	}

    "TMR_Signature_GRO_Hotel" {
	    $logo_src = "TMR_Signature_GRO_Hotel";
      	$sWeb = "www.hotels.com";
      	$sCompany = "ČR, a.s.";
		$sStreetAddress = "Ostr 75";
		$sPostalCode = "73914";
		$sStreet = "Oste";
		$sState = "Czech republic";
		break
	}

   

   	"TMR_Signature_NT_HotelTriStudnicky" {
      	$logo_src = "TMR_Signature_NT_HotelTriStudnicky";
      	$sWeb = "www.hot.com";
      	$sCompany = "Tatorts, a.s.";
		$sStreetAddress = "Demäna 72";
		$sPostalCode = "03101";
		$sStreet = "Lip Mik";
		$sState = "Slovakia";
		break
	}

    
    default {
        $sWeb = "www.sme.sk";
        $logo_src = "TMR_Signature_";
        $sCompany = "Tay , a.s.";
        $sStreetAddress = "Dea 72";
        $sPostalCode = "03101";
        $sStreet = "Liptuláš";
        $sState = "Slovakia";
        break
    }
}
}
Write-Host "$i. $Tab" -NoNewline

Write-Host $UserObj.Displayname $Tab $Tab $UserObj.Mail $Tab -NoNewline

if($sDepartment.length -gt 1) {
  $sPosition = "$sPosition  |  "  
}


Write-Host $sMobile "   " -NoNewline

Set-MailboxMessageConfiguration -Identity $UserObj.UserPrincipalName -AutoAddSignature $True -AutoAddSignatureOnMobile $True -AutoAddSignatureOnReply $false -SignatureHtml "
<html>
<head>
<meta content='text/html; charset=windows-1250'>
<meta content='MSHTML 6.00.2900.3059' name='GENERATOR'>
<style type='text/css'>
<!--
a.email {font-family:Arial;font-size:13px;font-weight:700;color:#363f43}
a.web {font-family:Arial;font-size:13px;font-weight:700;color:#d70424}
-->
</style>
</head>
<body>
<div><span>$($sPozdav)</span><br>
</div>
<br>
<table cellspacing='0' cellpadding='0' style='border-collapse:collapse'>
<tbody>
<tr>
<td colspan='2' class='meno' style='height:2px; font-family:Arial; font-size:16px; font-weight:700; color:rgb(215,4,36)'>$($sMeno) $($sPriezvisko)</td>
</tr>
<tr>
<td colspan='2' class='pozicia' style='height:2px; font-family:Arial; font-size:12px; font-weight:400; color:rgb(54,63,67)'>$($sPosition) $($sDepartment)</td>
</tr>
<tr>
<td colspan='2' class='odskok' style='height:2px; height:12px'></td>
</tr>
<tr>
<td width='18' class='i_telefon' style='height:2px;'><img src='' style='max-width:100%'><span id='dataURI' style='display:none'>data:image/png;base64,$([convert]::ToBase64String((Get-Content .\Telefon-x.png -Encoding byte)))</span></td>
<td class='telefon' style='height:2px; font-family:Arial; font-size:13px; font-weight:700; color:rgb(54,63,67)'>$($sMobile)</td>
</tr>
<tr>
<td width='18' class='i_email' style='height:2px;)'><img src='' style='max-width:100%'><span id='dataURI' style='display:none'>data:image/png;base64,$([convert]::ToBase64String((Get-Content .\Mail-x.png -Encoding byte)))</span></td>
<td style='height:2px'><a class='email' href='mailto:$($sEmail)' style='font-family:Arial; font-size:13px; font-weight:700; color:rgb
(54,63,67)'>$($sEmail)</a></td>
</tr>
<tr>
<td colspan='2' class='odskok' style='height:2px; height:12px'></td>
</tr>
<tr>
<td width='18' class='i_adresa' style='height:2px;vertical-align:top;'><img src='' style='max-width:100%'><span id='dataURI' style='display:none'>data:image/png;base64,$([convert]::ToBase64String((Get-Content .\Poloha-x.png -Encoding byte)))</span></td>
<td style='height:2px'><span class='firma' style='font-family:Arial; font-size:12px; font-weight:700; color:rgb(54,63,67)'>$($sCompany)</span><br>
<span class='adresa' style='font-family:Arial; font-size:12px; font-weight:400; color:rgb(54,63,67)'>$($sStreetAddress)</span><br>
<span class='pscmesto' style='font-family:Arial; font-size:12px; font-weight:400; color:rgb(54,63,67)'>$($sPostalCode) $($sStreet)</span><br>
<span class='stat' style='font-family:Arial; font-size:12px; font-weight:400; color:rgb(54,63,67)'>$($sState)</span>
</td>
</tr>
<tr>
<td colspan='2' class='odskok' style='height:2px; height:12px'></td>
</tr>
<tr>
<td width='18' class='i_web' style='height:2px;'><img src='' style='max-width:100%'><span id='dataURI' style='display:none'>data:image/png;base64,$([convert]::ToBase64String((Get-Content .\Web-x.png -Encoding byte)))</span></td>
<td style='height:2px'><a class='web' href='http://$($sWeb)' style='font-family:Arial; font-size:13px; font-weight:700; color:rgb(215,4,36)
'>$($sWeb)</a></td>
</tr>
<tr>
<td width='18' height='8' class='legend' style='height:2px; text-align:center; width:18px'>
</td>
<td style='height:2px'></td>
</tr>
<tr>
</tr>
</tbody>
</table>
<table cellspacing='0' cellpadding='0' style='border-collapse:collapse'>
<tbody>
<tr>
<td style='height:2px'><img src='' style='max-width:100%'><span id='dataURI' style='display:none'>data:image/png;base64,$([convert]::ToBase64String((Get-Content .\$logo_src.png -Encoding byte)))</span></td>
</tr>
</tbody>
</table>
</body>
</html>
"
Write-host " OK" -ForegroundColor Green
}
