def html_input(x)
  <<~HTML
    <html><head><title>blank</title>
    <meta name="Originator" content="Me"/>
    </head>
    <body>
    #{x}
    </body></html>
  HTML
end

def html_input_no_title(x)
  <<~HTML
    <html><head>
    <meta name="Originator" content="Me"/>
    </head>
    <body>
    #{x}
    </body></html>
  HTML
end

def html_input_empty_head(x)
  <<~HTML
    <html><head></head>
    <body>
    #{x}
    </body></html>
  HTML
end

WORD_HDR = <<~HDR
MIME-Version: 1.0
Content-Type: multipart/related; boundary="----=_NextPart_"

------=_NextPart_
Content-Location: file:///C:/Doc/test.htm
Content-Type: text/html; charset="utf-8"

<?xml version="1.0"?>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]>
<xml>
<w:WordDocument>
<w:View>Print</w:View>
<w:Zoom>100</w:Zoom>
<w:DoNotOptimizeForBrowser/>
</w:WordDocument>
</xml>
<![endif]-->
<meta http-equiv=Content-Type content="text/html; charset=utf-8"/>

  <link rel=File-List href="test_files/filelist.xml"/>
<title>blank</title><style><![CDATA[
  <!--
HDR

WORD_HDR_END = <<~HDR
-->
]]></style>
<meta name="Originator" content="Me"/>
</head>
HDR

def word_body(x, fn)
  <<~BODY
<body>
  #{x}
#{fn}</body></html>
  BODY
end

WORD_FTR1 = <<~FTR
  ------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: #{Html2Doc::mime_type('filelist.xml')}

PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
bGVsaXN0LnhtbCIvPgo8L3htbD4K

------=_NextPart_--
FTR

WORD_FTR2 = <<~FTR
  ------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: #{Html2Doc::mime_type('filelist.xml')}
PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9ImZp
bGVsaXN0LnhtbCIvPgogIDxvOkZpbGUgSFJlZj0iaGVhZGVyLmh0bWwiLz4KPC94bWw+Cg==
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/header.html
Content-Transfer-Encoding: base64
Content-Type: text/html charset="utf-8"
PGh0bWwgeG1sbnM6dj0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTp2bWwiDQp4bWxuczpvPSJ1
cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiDQp4bWxuczp3PSJ1cm46c2No
ZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTp3b3JkIg0KeG1sbnM6bT0iaHR0cDovL3NjaGVtYXMu
bWljcm9zb2Z0LmNvbS9vZmZpY2UvMjAwNC8xMi9vbW1sIg0KeG1sbnM6bXY9Imh0dHA6Ly9tYWNW
bWxTY2hlbWFVcmkiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy9UUi9SRUMtaHRtbDQwIj4NCg0K
PGhlYWQ+DQo8bWV0YSBuYW1lPVRpdGxlIGNvbnRlbnQ9IiI+DQo8bWV0YSBuYW1lPUtleXdvcmRz
IGNvbnRlbnQ9IiI+DQo8bWV0YSBodHRwLWVxdWl2PUNvbnRlbnQtVHlwZSBjb250ZW50PSJ0ZXh0
L2h0bWw7IGNoYXJzZXQ9dXRmLTgiPg0KPG1ldGEgbmFtZT1Qcm9nSWQgY29udGVudD1Xb3JkLkRv
Y3VtZW50Pg0KPG1ldGEgbmFtZT1HZW5lcmF0b3IgY29udGVudD0iTWljcm9zb2Z0IFdvcmQgMTUi
Pg0KPG1ldGEgbmFtZT1PcmlnaW5hdG9yIGNvbnRlbnQ9Ik1pY3Jvc29mdCBXb3JkIDE1Ij4NCjxs
aW5rIGlkPU1haW4tRmlsZSByZWw9TWFpbi1GaWxlIGhyZWY9IkZJTEVOQU1FLmh0bWwiPg0KPCEt
LVtpZiBndGUgbXNvIDldPjx4bWw+DQogPG86c2hhcGVkZWZhdWx0cyB2OmV4dD0iZWRpdCIgc3Bp
ZG1heD0iMjA0OSIvPg0KPC94bWw+PCFbZW5kaWZdLS0+DQo8L2hlYWQ+DQoNCjxib2R5IGxhbmc9
RU4gbGluaz1ibHVlIHZsaW5rPSIjOTU0RjcyIj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6
Zm9vdG5vdGUtc2VwYXJhdG9yJyBpZD1mcz4NCg0KPHAgY2xhc3M9TXNvTm9ybWFsIHN0eWxlPSdt
YXJnaW4tYm90dG9tOjBjbTttYXJnaW4tYm90dG9tOi4wMDAxcHQ7bGluZS1oZWlnaHQ6DQpub3Jt
YWwnPjxzcGFuIGxhbmc9RU4tR0I+PHNwYW4gc3R5bGU9J21zby1zcGVjaWFsLWNoYXJhY3Rlcjpm
b290bm90ZS1zZXBhcmF0b3InPjwhW2lmICFzdXBwb3J0Rm9vdG5vdGVzXT4NCg0KPGhyIGFsaWdu
PWxlZnQgc2l6ZT0xIHdpZHRoPSIzMyUiPg0KDQo8IVtlbmRpZl0+PC9zcGFuPjwvc3Bhbj48L3A+
DQoNCjwvZGl2Pg0KDQo8ZGl2IHN0eWxlPSdtc28tZWxlbWVudDpmb290bm90ZS1jb250aW51YXRp
b24tc2VwYXJhdG9yJyBpZD1mY3M+DQoNCjxwIGNsYXNzPU1zb05vcm1hbCBzdHlsZT0nbWFyZ2lu
LWJvdHRvbTowY207bWFyZ2luLWJvdHRvbTouMDAwMXB0O2xpbmUtaGVpZ2h0Og0Kbm9ybWFsJz48
c3BhbiBsYW5nPUVOLUdCPjxzcGFuIHN0eWxlPSdtc28tc3BlY2lhbC1jaGFyYWN0ZXI6Zm9vdG5v
dGUtY29udGludWF0aW9uLXNlcGFyYXRvcic+PCFbaWYgIXN1cHBvcnRGb290bm90ZXNdPg0KDQo8
aHIgYWxpZ249bGVmdCBzaXplPTE+DQoNCjwhW2VuZGlmXT48L3NwYW4+PC9zcGFuPjwvcD4NCg0K
PC9kaXY+DQoNCjxkaXYgc3R5bGU9J21zby1lbGVtZW50OmVuZG5vdGUtc2VwYXJhdG9yJyBpZD1l
cz4NCg0KPHAgY2xhc3M9TXNvTm9ybWFsIHN0eWxlPSdtYXJnaW4tYm90dG9tOjBjbTttYXJnaW4t
Ym90dG9tOi4wMDAxcHQ7bGluZS1oZWlnaHQ6DQpub3JtYWwnPjxzcGFuIGxhbmc9RU4tR0I+PHNw
YW4gc3R5bGU9J21zby1zcGVjaWFsLWNoYXJhY3Rlcjpmb290bm90ZS1zZXBhcmF0b3InPjwhW2lm
ICFzdXBwb3J0Rm9vdG5vdGVzXT4NCg0KPGhyIGFsaWduPWxlZnQgc2l6ZT0xIHdpZHRoPSIzMyUi
Pg0KDQo8IVtlbmRpZl0+PC9zcGFuPjwvc3Bhbj48L3A+DQoNCjwvZGl2Pg0KDQo8ZGl2IHN0eWxl
PSdtc28tZWxlbWVudDplbmRub3RlLWNvbnRpbnVhdGlvbi1zZXBhcmF0b3InIGlkPWVjcz4NCg0K
PHAgY2xhc3M9TXNvTm9ybWFsIHN0eWxlPSdtYXJnaW4tYm90dG9tOjBjbTttYXJnaW4tYm90dG9t
Oi4wMDAxcHQ7bGluZS1oZWlnaHQ6DQpub3JtYWwnPjxzcGFuIGxhbmc9RU4tR0I+PHNwYW4gc3R5
bGU9J21zby1zcGVjaWFsLWNoYXJhY3Rlcjpmb290bm90ZS1jb250aW51YXRpb24tc2VwYXJhdG9y
Jz48IVtpZiAhc3VwcG9ydEZvb3Rub3Rlc10+DQoNCjxociBhbGlnbj1sZWZ0IHNpemU9MT4NCg0K
PCFbZW5kaWZdPjwvc3Bhbj48L3NwYW4+PC9wPg0KDQo8L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNv
LWVsZW1lbnQ6aGVhZGVyJyBpZD1laDE+DQoNCjxwIGNsYXNzPU1zb0hlYWRlciBhbGlnbj1sZWZ0
IHN0eWxlPSd0ZXh0LWFsaWduOmxlZnQ7bGluZS1oZWlnaHQ6MTIuMHB0Ow0KbXNvLWxpbmUtaGVp
Z2h0LXJ1bGU6ZXhhY3RseSc+PHNwYW4gbGFuZz1FTi1HQj5JU08vSUVDJm5ic3A7Q0QgMTczMDEt
MToyMDE2KEUpPC9zcGFuPjwvcD4NCg0KPC9kaXY+DQoNCjxkaXYgc3R5bGU9J21zby1lbGVtZW50
OmhlYWRlcicgaWQ9aDE+DQoNCjxwIGNsYXNzPU1zb0hlYWRlciBzdHlsZT0nbWFyZ2luLWJvdHRv
bToxOC4wcHQnPjxzcGFuIGxhbmc9RU4tR0INCnN0eWxlPSdmb250LXNpemU6MTAuMHB0O21zby1i
aWRpLWZvbnQtc2l6ZToxMS4wcHQ7Zm9udC13ZWlnaHQ6bm9ybWFsJz7CqQ0KSVNPL0lFQyZuYnNw
OzIwMTYmbmJzcDvigJMgQWxsIHJpZ2h0cyByZXNlcnZlZDwvc3Bhbj48c3BhbiBsYW5nPUVOLUdC
DQpzdHlsZT0nZm9udC13ZWlnaHQ6bm9ybWFsJz48bzpwPjwvbzpwPjwvc3Bhbj48L3A+DQoNCjwv
ZGl2Pg0KDQo8ZGl2IHN0eWxlPSdtc28tZWxlbWVudDpmb290ZXInIGlkPWVmMT4NCg0KPHAgY2xh
c3M9TXNvRm9vdGVyIHN0eWxlPSdtYXJnaW4tdG9wOjEyLjBwdDtsaW5lLWhlaWdodDoxMi4wcHQ7
bXNvLWxpbmUtaGVpZ2h0LXJ1bGU6DQpleGFjdGx5Jz48IS0tW2lmIHN1cHBvcnRGaWVsZHNdPjxi
IHN0eWxlPSdtc28tYmlkaS1mb250LXdlaWdodDpub3JtYWwnPjxzcGFuDQpsYW5nPUVOLUdCIHN0
eWxlPSdmb250LXNpemU6MTAuMHB0O21zby1iaWRpLWZvbnQtc2l6ZToxMS4wcHQnPjxzcGFuDQpz
dHlsZT0nbXNvLWVsZW1lbnQ6ZmllbGQtYmVnaW4nPjwvc3Bhbj48c3Bhbg0Kc3R5bGU9J21zby1z
cGFjZXJ1bjp5ZXMnPsKgPC9zcGFuPlBBR0U8c3BhbiBzdHlsZT0nbXNvLXNwYWNlcnVuOnllcyc+
wqDCoA0KPC9zcGFuPlwqIE1FUkdFRk9STUFUIDxzcGFuIHN0eWxlPSdtc28tZWxlbWVudDpmaWVs
ZC1zZXBhcmF0b3InPjwvc3Bhbj48L3NwYW4+PC9iPjwhW2VuZGlmXS0tPjxiDQpzdHlsZT0nbXNv
LWJpZGktZm9udC13ZWlnaHQ6bm9ybWFsJz48c3BhbiBsYW5nPUVOLUdCIHN0eWxlPSdmb250LXNp
emU6MTAuMHB0Ow0KbXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+PHNwYW4gc3R5bGU9J21zby1u
by1wcm9vZjp5ZXMnPjI8L3NwYW4+PC9zcGFuPjwvYj48IS0tW2lmIHN1cHBvcnRGaWVsZHNdPjxi
DQpzdHlsZT0nbXNvLWJpZGktZm9udC13ZWlnaHQ6bm9ybWFsJz48c3BhbiBsYW5nPUVOLUdCIHN0
eWxlPSdmb250LXNpemU6MTAuMHB0Ow0KbXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+PHNwYW4g
c3R5bGU9J21zby1lbGVtZW50OmZpZWxkLWVuZCc+PC9zcGFuPjwvc3Bhbj48L2I+PCFbZW5kaWZd
LS0+PHNwYW4NCmxhbmc9RU4tR0Igc3R5bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7bXNvLWJpZGktZm9u
dC1zaXplOjExLjBwdCc+PHNwYW4NCnN0eWxlPSdtc28tdGFiLWNvdW50OjEnPsKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqAgPC9zcGFuPsKpDQpJ
U08vSUVDJm5ic3A7MjAxNiZuYnNwO+KAkyBBbGwgcmlnaHRzIHJlc2VydmVkPG86cD48L286cD48
L3NwYW4+PC9wPg0KDQo8L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6aGVhZGVyJyBp
ZD1laDI+DQoNCjxwIGNsYXNzPU1zb0hlYWRlciBhbGlnbj1sZWZ0IHN0eWxlPSd0ZXh0LWFsaWdu
OmxlZnQ7bGluZS1oZWlnaHQ6MTIuMHB0Ow0KbXNvLWxpbmUtaGVpZ2h0LXJ1bGU6ZXhhY3RseSc+
PHNwYW4gbGFuZz1FTi1HQj5JU08vSUVDJm5ic3A7Q0QgMTczMDEtMToyMDE2KEUpPC9zcGFuPjwv
cD4NCg0KPC9kaXY+DQoNCjxkaXYgc3R5bGU9J21zby1lbGVtZW50OmhlYWRlcicgaWQ9aDI+DQoN
CjxwIGNsYXNzPU1zb0hlYWRlciBhbGlnbj1yaWdodCBzdHlsZT0ndGV4dC1hbGlnbjpyaWdodDts
aW5lLWhlaWdodDoxMi4wcHQ7DQptc28tbGluZS1oZWlnaHQtcnVsZTpleGFjdGx5Jz48c3BhbiBs
YW5nPUVOLUdCPklTTy9JRUMmbmJzcDtDRCAxNzMwMS0xOjIwMTYoRSk8L3NwYW4+PC9wPg0KDQo8
L2Rpdj4NCg0KPGRpdiBzdHlsZT0nbXNvLWVsZW1lbnQ6Zm9vdGVyJyBpZD1lZjI+DQoNCjxwIGNs
YXNzPU1zb0Zvb3RlciBzdHlsZT0nbGluZS1oZWlnaHQ6MTIuMHB0O21zby1saW5lLWhlaWdodC1y
dWxlOmV4YWN0bHknPjwhLS1baWYgc3VwcG9ydEZpZWxkc10+PHNwYW4NCmxhbmc9RU4tR0Igc3R5
bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7bXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+PHNwYW4NCnN0
eWxlPSdtc28tZWxlbWVudDpmaWVsZC1iZWdpbic+PC9zcGFuPjxzcGFuDQpzdHlsZT0nbXNvLXNw
YWNlcnVuOnllcyc+wqA8L3NwYW4+UEFHRTxzcGFuIHN0eWxlPSdtc28tc3BhY2VydW46eWVzJz7C
oMKgDQo8L3NwYW4+XCogTUVSR0VGT1JNQVQgPHNwYW4gc3R5bGU9J21zby1lbGVtZW50OmZpZWxk
LXNlcGFyYXRvcic+PC9zcGFuPjwvc3Bhbj48IVtlbmRpZl0tLT48c3Bhbg0KbGFuZz1FTi1HQiBz
dHlsZT0nZm9udC1zaXplOjEwLjBwdDttc28tYmlkaS1mb250LXNpemU6MTEuMHB0Jz48c3Bhbg0K
c3R5bGU9J21zby1uby1wcm9vZjp5ZXMnPmlpPC9zcGFuPjwvc3Bhbj48IS0tW2lmIHN1cHBvcnRG
aWVsZHNdPjxzcGFuDQpsYW5nPUVOLUdCIHN0eWxlPSdmb250LXNpemU6MTAuMHB0O21zby1iaWRp
LWZvbnQtc2l6ZToxMS4wcHQnPjxzcGFuDQpzdHlsZT0nbXNvLWVsZW1lbnQ6ZmllbGQtZW5kJz48
L3NwYW4+PC9zcGFuPjwhW2VuZGlmXS0tPjxzcGFuIGxhbmc9RU4tR0INCnN0eWxlPSdmb250LXNp
emU6MTAuMHB0O21zby1iaWRpLWZvbnQtc2l6ZToxMS4wcHQnPjxzcGFuIHN0eWxlPSdtc28tdGFi
LWNvdW50Og0KMSc+wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoCA8L3NwYW4+wqkNCklTTy9JRUMmbmJzcDsyMDE2Jm5ic3A74oCTIEFsbCByaWdo
dHMgcmVzZXJ2ZWQ8bzpwPjwvbzpwPjwvc3Bhbj48L3A+DQoNCjwvZGl2Pg0KDQo8ZGl2IHN0eWxl
PSdtc28tZWxlbWVudDpmb290ZXInIGlkPWYyPg0KDQo8cCBjbGFzcz1Nc29Gb290ZXIgc3R5bGU9
J2xpbmUtaGVpZ2h0OjEyLjBwdCc+PHNwYW4gbGFuZz1FTi1HQg0Kc3R5bGU9J2ZvbnQtc2l6ZTox
MC4wcHQ7bXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+wqkgSVNPL0lFQyZuYnNwOzIwMTYmbmJz
cDvigJMgQWxsDQpyaWdodHMgcmVzZXJ2ZWQ8c3BhbiBzdHlsZT0nbXNvLXRhYi1jb3VudDoxJz7C
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoCA8L3Nw
YW4+PC9zcGFuPjwhLS1baWYgc3VwcG9ydEZpZWxkc10+PHNwYW4NCmxhbmc9RU4tR0Igc3R5bGU9
J2ZvbnQtc2l6ZToxMC4wcHQ7bXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+PHNwYW4NCnN0eWxl
PSdtc28tZWxlbWVudDpmaWVsZC1iZWdpbic+PC9zcGFuPiBQQUdFPHNwYW4gc3R5bGU9J21zby1z
cGFjZXJ1bjp5ZXMnPsKgwqANCjwvc3Bhbj5cKiBNRVJHRUZPUk1BVCA8c3BhbiBzdHlsZT0nbXNv
LWVsZW1lbnQ6ZmllbGQtc2VwYXJhdG9yJz48L3NwYW4+PC9zcGFuPjwhW2VuZGlmXS0tPjxzcGFu
DQpsYW5nPUVOLUdCIHN0eWxlPSdmb250LXNpemU6MTAuMHB0O21zby1iaWRpLWZvbnQtc2l6ZTox
MS4wcHQnPjxzcGFuDQpzdHlsZT0nbXNvLW5vLXByb29mOnllcyc+aWlpPC9zcGFuPjwvc3Bhbj48
IS0tW2lmIHN1cHBvcnRGaWVsZHNdPjxzcGFuDQpsYW5nPUVOLUdCIHN0eWxlPSdmb250LXNpemU6
MTAuMHB0O21zby1iaWRpLWZvbnQtc2l6ZToxMS4wcHQnPjxzcGFuDQpzdHlsZT0nbXNvLWVsZW1l
bnQ6ZmllbGQtZW5kJz48L3NwYW4+PC9zcGFuPjwhW2VuZGlmXS0tPjxzcGFuIGxhbmc9RU4tR0IN
CnN0eWxlPSdmb250LXNpemU6MTAuMHB0O21zby1iaWRpLWZvbnQtc2l6ZToxMS4wcHQnPjxvOnA+
PC9vOnA+PC9zcGFuPjwvcD4NCg0KPC9kaXY+DQoNCjxkaXYgc3R5bGU9J21zby1lbGVtZW50OmZv
b3RlcicgaWQ9ZWYzPg0KDQo8cCBjbGFzcz1Nc29Gb290ZXIgc3R5bGU9J21hcmdpbi10b3A6MTIu
MHB0O2xpbmUtaGVpZ2h0OjEyLjBwdDttc28tbGluZS1oZWlnaHQtcnVsZToNCmV4YWN0bHknPjwh
LS1baWYgc3VwcG9ydEZpZWxkc10+PGIgc3R5bGU9J21zby1iaWRpLWZvbnQtd2VpZ2h0Om5vcm1h
bCc+PHNwYW4NCmxhbmc9RU4tR0Igc3R5bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7bXNvLWJpZGktZm9u
dC1zaXplOjExLjBwdCc+PHNwYW4NCnN0eWxlPSdtc28tZWxlbWVudDpmaWVsZC1iZWdpbic+PC9z
cGFuPjxzcGFuDQpzdHlsZT0nbXNvLXNwYWNlcnVuOnllcyc+wqA8L3NwYW4+UEFHRTxzcGFuIHN0
eWxlPSdtc28tc3BhY2VydW46eWVzJz7CoMKgDQo8L3NwYW4+XCogTUVSR0VGT1JNQVQgPHNwYW4g
c3R5bGU9J21zby1lbGVtZW50OmZpZWxkLXNlcGFyYXRvcic+PC9zcGFuPjwvc3Bhbj48L2I+PCFb
ZW5kaWZdLS0+PGINCnN0eWxlPSdtc28tYmlkaS1mb250LXdlaWdodDpub3JtYWwnPjxzcGFuIGxh
bmc9RU4tR0Igc3R5bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7DQptc28tYmlkaS1mb250LXNpemU6MTEu
MHB0Jz48c3BhbiBzdHlsZT0nbXNvLW5vLXByb29mOnllcyc+Mjwvc3Bhbj48L3NwYW4+PC9iPjwh
LS1baWYgc3VwcG9ydEZpZWxkc10+PGINCnN0eWxlPSdtc28tYmlkaS1mb250LXdlaWdodDpub3Jt
YWwnPjxzcGFuIGxhbmc9RU4tR0Igc3R5bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7DQptc28tYmlkaS1m
b250LXNpemU6MTEuMHB0Jz48c3BhbiBzdHlsZT0nbXNvLWVsZW1lbnQ6ZmllbGQtZW5kJz48L3Nw
YW4+PC9zcGFuPjwvYj48IVtlbmRpZl0tLT48c3Bhbg0KbGFuZz1FTi1HQiBzdHlsZT0nZm9udC1z
aXplOjEwLjBwdDttc28tYmlkaS1mb250LXNpemU6MTEuMHB0Jz48c3Bhbg0Kc3R5bGU9J21zby10
YWItY291bnQ6MSc+wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoCA8L3NwYW4+wqkNCklTTy9JRUMmbmJzcDsyMDE2Jm5ic3A74oCTIEFsbCByaWdo
dHMgcmVzZXJ2ZWQ8bzpwPjwvbzpwPjwvc3Bhbj48L3A+DQoNCjwvZGl2Pg0KDQo8ZGl2IHN0eWxl
PSdtc28tZWxlbWVudDpmb290ZXInIGlkPWYzPg0KDQo8cCBjbGFzcz1Nc29Gb290ZXIgc3R5bGU9
J2xpbmUtaGVpZ2h0OjEyLjBwdCc+PHNwYW4gbGFuZz1FTi1HQg0Kc3R5bGU9J2ZvbnQtc2l6ZTox
MC4wcHQ7bXNvLWJpZGktZm9udC1zaXplOjExLjBwdCc+wqkgSVNPL0lFQyZuYnNwOzIwMTYmbmJz
cDvigJMgQWxsDQpyaWdodHMgcmVzZXJ2ZWQ8c3BhbiBzdHlsZT0nbXNvLXRhYi1jb3VudDoxJz7C
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDC
oMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKg
wqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgwqDCoMKgIDwv
c3Bhbj48L3NwYW4+PCEtLVtpZiBzdXBwb3J0RmllbGRzXT48Yg0Kc3R5bGU9J21zby1iaWRpLWZv
bnQtd2VpZ2h0Om5vcm1hbCc+PHNwYW4gbGFuZz1FTi1HQiBzdHlsZT0nZm9udC1zaXplOjEwLjBw
dDsNCm1zby1iaWRpLWZvbnQtc2l6ZToxMS4wcHQnPjxzcGFuIHN0eWxlPSdtc28tZWxlbWVudDpm
aWVsZC1iZWdpbic+PC9zcGFuPg0KUEFHRTxzcGFuIHN0eWxlPSdtc28tc3BhY2VydW46eWVzJz7C
oMKgIDwvc3Bhbj5cKiBNRVJHRUZPUk1BVCA8c3Bhbg0Kc3R5bGU9J21zby1lbGVtZW50OmZpZWxk
LXNlcGFyYXRvcic+PC9zcGFuPjwvc3Bhbj48L2I+PCFbZW5kaWZdLS0+PGINCnN0eWxlPSdtc28t
YmlkaS1mb250LXdlaWdodDpub3JtYWwnPjxzcGFuIGxhbmc9RU4tR0Igc3R5bGU9J2ZvbnQtc2l6
ZToxMC4wcHQ7DQptc28tYmlkaS1mb250LXNpemU6MTEuMHB0Jz48c3BhbiBzdHlsZT0nbXNvLW5v
LXByb29mOnllcyc+Mzwvc3Bhbj48L3NwYW4+PC9iPjwhLS1baWYgc3VwcG9ydEZpZWxkc10+PGIN
CnN0eWxlPSdtc28tYmlkaS1mb250LXdlaWdodDpub3JtYWwnPjxzcGFuIGxhbmc9RU4tR0Igc3R5
bGU9J2ZvbnQtc2l6ZToxMC4wcHQ7DQptc28tYmlkaS1mb250LXNpemU6MTEuMHB0Jz48c3BhbiBz
dHlsZT0nbXNvLWVsZW1lbnQ6ZmllbGQtZW5kJz48L3NwYW4+PC9zcGFuPjwvYj48IVtlbmRpZl0t
LT48c3Bhbg0KbGFuZz1FTi1HQiBzdHlsZT0nZm9udC1zaXplOjEwLjBwdDttc28tYmlkaS1mb250
LXNpemU6MTEuMHB0Jz48bzpwPjwvbzpwPjwvc3Bhbj48L3A+DQoNCjwvZGl2Pg0KDQo8L2JvZHk+
DQoNCjwvaHRtbD4NCg==

------=_NextPart_--
FTR

WORD_FTR3 = <<~FTR
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/filelist.xml
Content-Transfer-Encoding: base64
Content-Type: #{Html2Doc::mime_type('filelist.xml')}

PHhtbCB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiPgog
ICAgICAgIDxvOk1haW5GaWxlIEhSZWY9Ii4uL3Rlc3QuaHRtIi8+ICA8bzpGaWxlIEhSZWY9IjFh
YzIwNjVmLTAzZjAtNGM3YS1iOWE2LTkyZTgyMDU5MWJmMC5wbmciLz4KICA8bzpGaWxlIEhSZWY9
ImZpbGVsaXN0LnhtbCIvPgo8L3htbD4K
------=_NextPart_
Content-Location: file:///C:/Doc/test_files/cb7b0d19-891e-4634-815a-570d019d454c.png
Content-Transfer-Encoding: base64
Content-Type: image/png
------=_NextPart_--
FTR

DEFAULT_STYLESHEET = File.read("lib/html2doc/wordstyle.css", encoding: "utf-8").freeze

def guid_clean(x)
  x.gsub(/NextPart_[0-9a-f.]+/, "NextPart_")
end

def image_clean(x)
  x.gsub(%r{[0-9a-f-]+\.png}, "image.png").
    gsub(%r{[0-9a-f-]+\.gif}, "image.gif").
    gsub(%r{[0-9a-f-]+\.(jpeg|jpg)}, "image.jpg").
    gsub(%r{------=_NextPart_\s+Content-Location: file:///C:/Doc/test_files/image\.(png|gif).*?\s-----=_NextPart_}m, "------=_NextPart_").
    gsub(%r{Content-Type: image/(png|gif|jpeg)[^-]*------=_NextPart_-?-?}m, "").
    gsub(%r{ICAgICAg[^-]*-----}m, "-----").
    gsub(%r{\s*</img>\s*}m, "</img>").
    gsub(%r{</body>\s*</html>}m, "</body></html>")
end

RSpec.describe Html2Doc do
  it "has a version number" do
    expect(Html2Doc::VERSION).not_to be nil
  end

  it "processes a blank document" do
    Html2Doc.process(html_input(""), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "removes any temp files" do
    File.delete("test.doc")
    Html2Doc.process(html_input(""), filename: "test")
    expect(File.exist?("test.doc")).to be true
    expect(File.exist?("test.htm")).to be false
    expect(File.exist?("test_files")).to be false
  end

  it "processes a stylesheet in an HTML document with a title" do
    Html2Doc.process(html_input(""), filename: "test", stylesheet: "lib/html2doc/wordstyle.css")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a stylesheet in an HTML document without a title" do
    Html2Doc.process(html_input_no_title(""), filename: "test", stylesheet: "lib/html2doc/wordstyle.css")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR.sub("<title>blank</title>", "")} 
    #{DEFAULT_STYLESHEET} #{WORD_HDR_END} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a stylesheet in an HTML document with an empty head" do
    Html2Doc.process(html_input_empty_head(""), filename: "test", stylesheet: "lib/html2doc/wordstyle.css")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR.sub("<title>blank</title>", "")}
    #{DEFAULT_STYLESHEET} 
    #{WORD_HDR_END.sub('<meta name="Originator" content="Me"/>'+"\n", "").sub("</style>\n</head>", "</style></head>")} 
    #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR1}
    OUTPUT
  end

  it "processes a header" do
    Html2Doc.process(html_input(""), filename: "test", header_file: "spec/header.html")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET.gsub(/FILENAME/, "test")} 
    #{WORD_HDR_END} #{word_body("", '<div style="mso-element:footnote-list"/>')} #{WORD_FTR2}
    OUTPUT
  end

  it "processes a header with an image" do
    Html2Doc.process(html_input(""), filename: "test", header_file: "spec/header_img.html")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).to match(%r{Content-Type: image/png})
  end


  it "processes a populated document" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very simple document</div>"
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body, '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "processes AsciiMath" do
    Html2Doc.process(html_input(%[<div>{{sum_(i=1)^n i^3=((n(n+1))/2)^2 text("integer"))}}</div>]), filename: "test", asciimathdelims: ["{{", "}}"])
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('
       <div><m:oMath>
       <m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="on"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub>
       <m:r><m:t>i=1</m:t></m:r>
       </m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSup></m:e></m:nary>
       <m:r><m:t>=</m:t></m:r>
       <m:sSup><m:e>
       <m:r><m:t>(</m:t></m:r>
       <m:f><m:fPr><m:type m:val="bar"></m:type></m:fPr><m:num>
       <m:r><m:t>n</m:t></m:r>
       <m:r><m:t>(n+1)</m:t></m:r>
       </m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
       <m:r><m:t>)</m:t></m:r>
       </m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
       <m:r><m:rPr><m:nor></m:nor></m:rPr><m:t>"integer"</m:t></m:r>
       </m:oMath>
    </div>', '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "left-aligns AsciiMath" do
    Html2Doc.process(html_input("<div style='text-align:left;'>{{sum_(i=1)^n i^3=((n(n+1))/2)^2}}</div>"), filename: "test", asciimathdelims: ["{{", "}}"])
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('
       <div style="text-align:left;"><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
       <m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="on"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub>
       <m:r><m:t>i=1</m:t></m:r>
       </m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSup></m:e></m:nary>
       <m:r><m:t>=</m:t></m:r>
       <m:sSup><m:e>
       <m:r><m:t>(</m:t></m:r>
       <m:f><m:fPr><m:type m:val="bar"></m:type></m:fPr><m:num>
       <m:r><m:t>n</m:t></m:r>
       <m:r><m:t>(n+1)</m:t></m:r>
       </m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
       <m:r><m:t>)</m:t></m:r>
       </m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
       </m:oMath>
       </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "right-aligns AsciiMath" do
    Html2Doc.process(html_input("<div style='text-align:right;'>{{sum_(i=1)^n i^3=((n(n+1))/2)^2}}</div>"), filename: "test", asciimathdelims: ["{{", "}}"])
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('
       <div style="text-align:right;"><m:oMathPara><m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr><m:oMath>
       <m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="on"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub>
       <m:r><m:t>i=1</m:t></m:r>
       </m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSup></m:e></m:nary>
       <m:r><m:t>=</m:t></m:r>
       <m:sSup><m:e>
       <m:r><m:t>(</m:t></m:r>
       <m:f><m:fPr><m:type m:val="bar"></m:type></m:fPr><m:num>
       <m:r><m:t>n</m:t></m:r>
       <m:r><m:t>(n+1)</m:t></m:r>
       </m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
       <m:r><m:t>)</m:t></m:r>
       </m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
       </m:oMath>
       </m:oMathPara></div>', '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "wraps msup after munderover in MathML" do
    Html2Doc.process(html_input("<div><math xmlns='http://www.w3.org/1998/Math/MathML'>
<munderover><mo>&#x2211;</mo><mrow><mi>i</mi><mo>=</mo><mn>0</mn></mrow><mrow><mi>n</mi></mrow></munderover><msup><mn>2</mn><mrow><mi>i</mi></mrow></msup></math></div>"), filename: "test", asciimathdelims: ["{{", "}}"])
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div><m:oMath>
       <m:nary><m:naryPr><m:chr m:val="&#x2211;"></m:chr><m:limLoc m:val="undOvr"></m:limLoc><m:grow m:val="on"></m:grow><m:subHide m:val="off"></m:subHide><m:supHide m:val="off"></m:supHide></m:naryPr><m:sub><m:r><m:t>i=0</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>2</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup></m:e></m:nary></m:oMath>
    </div>', '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "processes tabs" do
    simple_body = "<h1>Hello word!</h1>
    <div>This is a very &tab; simple document</div>"
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/\&tab;/, %[<span style="mso-tab-count:1">&#xA0; </span>]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "makes unstyled paragraphs be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <p>This is a very simple document</p>
    <p class="x">This style stays</p>'
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/<p>/, %[<p class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "makes unstyled list entries be MsoNormal" do
    simple_body = '<h1>Hello word!</h1>
    <ul>
    <li>This is a very simple document</li>
    <li class="x">This style stays</li>
    </ul>'
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body(simple_body.gsub(/<li>/, %[<li class="MsoNormal">]), '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "resizes images for height, in a file in a subdirectory" do
    simple_body = '<img src="19160-6.png">'
    Html2Doc.process(html_input(simple_body), filename: "spec/test")
    testdoc = File.read("spec/test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/png})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.png" width="400" height="388"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3)}
    OUTPUT
  end

  it "resizes images for width" do
    simple_body = '<img src="spec/19160-7.gif">'
    Html2Doc.process(html_input(simple_body), filename: "test")
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/gif})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.gif" width="400" height="118"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3).gsub(/image\.png/, "image.gif")}
    OUTPUT
  end

  it "resizes images for height" do
    simple_body = '<img src="spec/19160-8.jpg">'
    Html2Doc.process(html_input(simple_body), filename: "test")
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(testdoc).to match(%r{Content-Type: image/jpeg})
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="test_files/cb7b0d19-891e-4634-815a-570d019d454c.jpg" width="208" height="680"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR3).gsub(/image\.png/, "image.jpg")}
    OUTPUT
  end

  it "resizes images with missing or auto sizes" do
    image = { "src" => "spec/19160-8.jpg" }
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
    image["width"] = "20"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [20, 65]
    image.delete("width")
    image["height"] = "50"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [15, 50]
    image.delete("height")
    image["width"] = "500"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
    image.delete("width")
    image["height"] = "500"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
    image["width"] = "20"
    image["height"] = "auto"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [20, 65]
    image["width"] = "auto"
    image["height"] = "50"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [15, 50]
    image["width"] = "500"
    image["height"] = "auto"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
    image["width"] = "auto"
    image["height"] = "500"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
    image["width"] = "auto"
    image["height"] = "auto"
    expect(Html2Doc.image_resize(image, "spec/19160-8.jpg", 100, 100)).to eq [30, 100]
  end

  it "does not move images if they are external URLs" do
    simple_body = '<img src="https://example.com/19160-6.png">'
    Html2Doc.process(html_input(simple_body), filename: "test")
    testdoc = File.read("test.doc", encoding: "utf-8")
    expect(image_clean(guid_clean(testdoc))).to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{image_clean(word_body('<img src="https://example.com/19160-6.png"></img>', '<div style="mso-element:footnote-list"/>'))}
    #{image_clean(WORD_FTR1)}
    OUTPUT
  end

  it "warns about SVG" do
    simple_body = '<img src="https://example.com/19160-6.svg">'
    expect{ Html2Doc.process(html_input(simple_body), filename: "test") }.to output("https://example.com/19160-6.svg: SVG not supported\n").to_stderr
  end

  it "processes epub:type footnotes" do
    simple_body = '<div>This is a very simple 
     document<a epub:type="footnote" href="#a1">1</a> allegedly<a epub:type="footnote" href="#a2">2</a></div>
     <aside id="a1">Footnote</aside>
     <aside id="a2">Other Footnote</aside>'
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>This is a very simple
    document<a epub:type="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a epub:type="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
    '<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "processes class footnotes" do
    simple_body = '<div>This is a very simple
     document<a class="footnote" href="#a1">1</a> allegedly<a class="footnote" href="#a2">2</a></div>
     <aside id="a1">Footnote</aside>
     <aside id="a2">Other Footnote</aside>'
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>This is a very simple
    document<a class="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a class="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
    '<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p id="" class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "extracts paragraphs from footnotes" do
    simple_body = '<div>This is a very simple
     document<a class="footnote" href="#a1">1</a> allegedly<a class="footnote" href="#a2">2</a></div>
     <aside id="a1"><p>Footnote</p></aside>
     <div id="a2"><p>Other Footnote</p></div>'
    Html2Doc.process(html_input(simple_body), filename: "test")
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>This is a very simple
    document<a class="footnote" href="#_ftn1" style="mso-footnote-id:ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> allegedly<a class="footnote" href="#_ftn2" style="mso-footnote-id:ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a></div>',
    '<div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1">
<p class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1" title="" id="_ftnref1"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Footnote</p></div>
<div style="mso-element:footnote" id="ftn2">
<p class="MsoFootnoteText"><a style="mso-footnote-id:ftn2" href="#_ftn2" name="_ftnref2" title="" id="_ftnref2"><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>Other Footnote</p></div>
</div>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "labels lists with list styles" do
    simple_body = <<~BODY
      <div><ul>
      <li><div><p><ol><li><ul><li><p><ol><li><ol><li>A</li><li><p>B</p><p>B2</p></li><li>C</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ul></div>
    BODY
    Html2Doc.process(html_input(simple_body), filename: "test", liststyles: {ul: "l1", ol: "l2"})
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>
    <p style="mso-list:l1 level1 lfo1;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpFirst">A</p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpMiddle">B<p class="MsoListParagraphCxSpMiddle">B2</p></p><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpLast">C</p></p></p></p></div></p></div>',
    '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end


  it "restarts numbering of lists with list styles" do
    simple_body = <<~BODY
      <div>
      <ol><li><div><p><ol><li><ul><li><p><ol><li><ol><li>A</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ol>
      <ol><li><div><p><ol><li><ul><li><p><ol><li><ol><li>A</li></ol></li></ol></p></li></ul></li></ol></p></div></li></ol></div>
    BODY
    Html2Doc.process(html_input(simple_body), filename: "test", liststyles: {ul: "l1", ol: "l2"})
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>
      <p style="mso-list:l2 level1 lfo1;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo1;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo1;" class="MsoListParagraphCxSpFirst">A</p></p></p></p></div></p>
      <p style="mso-list:l2 level1 lfo2;" class="MsoListParagraphCxSpFirst"><div><p class="MsoNormal"><p style="mso-list:l2 level2 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level4 lfo2;" class="MsoListParagraphCxSpFirst"><p style="mso-list:l2 level5 lfo2;" class="MsoListParagraphCxSpFirst">A</p></p></p></p></div></p></div>',
    '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

  it "replaces id attributes with explicit a@name bookmarks" do
    simple_body = <<~BODY
      <div>
        <p id="a">Hello</p>
        <p id="b"/>
      </div>
    BODY
    Html2Doc.process(html_input(simple_body), filename: "test", liststyles: {ul: "l1", ol: "l2"})
    expect(guid_clean(File.read("test.doc", encoding: "utf-8"))).
      to match_fuzzy(<<~OUTPUT)
    #{WORD_HDR} #{DEFAULT_STYLESHEET} #{WORD_HDR_END}
    #{word_body('<div>
        <p class="MsoNormal"><a name="a" id="a"></a>Hello</p>
        <p class="MsoNormal"><a name="b" id="b"></a></p>
      </div>',
    '<div style="mso-element:footnote-list"/>')}
    #{WORD_FTR1}
    OUTPUT
  end

end
