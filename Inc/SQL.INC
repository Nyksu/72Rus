<%

var usrpsw="Update USERS Set PSW='%PS' Where ID=%ID"

var usrpsw="Update USERS Set PSW='%PS' Where ID=%ID"

var insurlhost="Insert into HOST_URL (ID,NAME,EMAIL,PHONE,NIKNAME,PSW,KEYWORD,STATE,CITY_ID,AREA_ID,COUNTRY_ID,COMPANY,DATE_REG) "+
		"values (%ID,'%NAME','%EMAIL','%PHONE','%NIK','%PS','%CW',1,%CITY,%AREA,%LAND,'%COMP','TODAY')"

var edurlhost="Update HOST_URL Set NAME='%NAME', EMAIL='%EMAIL', PHONE='%PHONE', CITY_ID=%CITY, AREA_ID=%AREA, COUNTRY_ID=%LAND, COMPANY='%COMP' where ID=%ID"

var usrpsw="Update USERS Set PSW='%PS' Where ID=%ID"

var inspubusr="Insert into smi_usr(ID,NAME,NIK_NAME,PSW,SMI_ID,USR_TYPE_ID,STATE,EMAIL,PHONE) values (%ID, '%NAM', '%NIK', '%PS', %SMI, %TP, 0, '%EML', '%PHN')"

var insheading="Insert into heading(ID,SMI_ID,NAME,HI_ID,PICTURE,URL,PERIOD,PAGE_LENGTH,SMI_USR_ID,ISNEWS) "+
		"values(%ID,%SMI,'%NAM',%HID, '', '%URL', %PER, %PL, %USR, %ISNEWS)"

var edheading="Update heading set NAME='%NAM', URL='%URL', PERIOD=%PER, PAGE_LENGTH=%PL, SMI_USR_ID=%USR, ISNEWS=%ISNEWS where ID=%ID"

var inspub="Insert into publication(ID,NAME,DIGEST,DATE_CREATE,PUBLIC_DATE,STATE,HEADING_ID,COMENT,AUTOR,URL,CREATER_ID,ISHTML) "+
		"values(%ID,'%NAM','%DIG','NOW','%PDAT',%ST, %HID, '%COM', '%AUT', '%URL', %USR, %ISH)"

var insblock="Insert into block_news(ID,NAME,SUBJ,SMI_ID) values (%ID,'%NAM','%SUBJ',%SMI)"

var edpub="Update publication set NAME='%NAM',DIGEST='%DIG',PUBLIC_DATE='%PDAT',COMENT='%COM',AUTOR='%AUT',URL='%URL',ISHTML=%ISH where id=%ID"

var edblock="Update block_news set NAME='%NAM', SUBJ='%SUBJ' where ID=%ID and SMI_ID=%SMI"

var allpubheading="select distinct t1.id from publication t1, getbranchheading(%HID,%SMI) p1 where p1.hid=t1.heading_id Union Select distinct t2.id from publication t2 where t2.heading_id=%HID1"


%>