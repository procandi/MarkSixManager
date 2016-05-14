# encoding: UTF-8


class SettingHandle
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