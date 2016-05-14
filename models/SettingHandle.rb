# encoding: UTF-8


class SettingHandle
=begin

recordset.Fields.each(){|field|
  p field.name
}

product
"PID"
"PName"
pid,pname

price
"SwiftCode"
"CID"
"PID"
"CurrentDate"
"CurrentPrice"
"WinningPrice"
"Upset"
swiftcode,cid,pid,currentdate,currentprice,winningprice,upset

custom
"CID"
"CName"
"CType"
"Address"
"OpenDate"
"BankID"
"Proportion"
"BonusTarget"
"Phone1"
"Phone2"
"Phone3"
"Phone4"
"Phone5"
"Phone6"
"Note"
cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note

[order]
"SwiftCode"
"CID"
"PID"
"CurrentDate"
"CurrentCount"
"WinningCount"
"AddMoney"
"BonusMoney"
"Note"
"Group"
swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group

=end


#=begin
  #環境基本資料
  REMOTE_DBTYPE="ORA"  #MSSQL or DB2 or ORA
  REMOTE_DBIP="192.168.167.98"
  REMOTE_DBID="system"
  REMOTE_DBPW="CSMH1"
  REMOTE_DBSID="PEDCV"
  LOCAL_DBTYPE="ORA"  #MSSQL or DB2 or ORA
  LOCAL_DBIP="192.168.110.123"
  LOCAL_DBID="system"
  LOCAL_DBPW="CSMH1"
  LOCAL_DBSID="COLON"
  DBTYPE="ORA"  #MSSQL or DB2 or ORA
  DBIP="192.168.110.123"
  DBID="system"
  DBPW="CSMH1"
  DBSID="CRIS"
  QUERYWHERE="(hisup='50' or hisup='51') and uni_key like '14%' and not (ascii(lower(substr(chartno,1,1)))>96 and ascii(lower(substr(chartno,1,1)))<123) and status='已報告' "
  
  #設定客製化呈現方式
  GET_DR_FROM_O=""    #是否當為門診時，顯示特定資料
  SET_DR_FROM_O=""
#=end

  
=begin
  #測試環境
  REMOTE_DBTYPE="ORA"  #MSSQL or DB2 or ORA
  REMOTE_DBIP="127.0.0.1"
  REMOTE_DBID="system"
  REMOTE_DBPW="CSMH1"
  REMOTE_DBSID="ORCL"
  LOCAL_DBTYPE="ORA"  #MSSQL or DB2 or ORA
  LOCAL_DBIP="127.0.0.1"
  LOCAL_DBID="system"
  LOCAL_DBPW="CSMH1"
  LOCAL_DBSID="ORCL"
  DBTYPE="ORA"  #MSSQL or DB2 or ORA
  DBIP="127.0.0.1"
  DBID="system"
  DBPW="CSMH1"
  DBSID="ORCL"
  QUERYWHERE="(hisup='50' or hisup='51') and uni_key like '14%' and not (ascii(lower(substr(chartno,1,1)))>96 and ascii(lower(substr(chartno,1,1)))<123) and status='已報告' "
    
  #設定客製化呈現方式
  GET_DR_FROM_O=""    #是否當為門診時，顯示特定資料
  SET_DR_FROM_O=""
=end
end