# -*- coding: utf-8 -*-
require 'win32ole'  
require 'spreadsheet'




#客戶每日交易明細
def CustomDailyTransactionDetail(connection,current_pid,current_cid,current_date)
	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
 	if(current_pid.to_i()>=100)
    	sql = "select * from product where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} order by CLng(PID);"
	else
		sql = "select * from product where PID='#{current_pid}';"
	end
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom where CID='#{current_cid}';"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from [order] where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CID='#{current_cid}' and CurrentDate='#{current_date}';"
	else
		sql = "select * from [order] where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate='#{current_date}';"
	end
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select top 1 * from price where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CID='#{current_cid}' and CurrentDate<='#{current_date}' order by CurrentDate desc;"
	else
		sql = "select top 1 * from price where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate<='#{current_date}' order by CurrentDate desc;"
	end
	recordset_price.Open(sql, connection)


	#預存出所有會需要列出的資料
	data_product = recordset_product.GetRows.transpose
=begin
	data_product.each(){|pid,pname|
		p pid,pname
	}
=end

	#data_custom = recordset_custom.GetRows.transpose
=begin
	data_custom.each(){|cid,cname|
		p cid,cname
	}
=end

	data_order = recordset_order.GetRows.transpose
=begin
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group|
		p swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group
	}
=end

	#data_price = recordset_price.GetRows.transpose
=begin
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		p swiftcode,cid,pid,currentdate,currentprice,winningprice,upset

	}
=end
	

	#Create a new Workbook
	book = Spreadsheet::Workbook.new

	#Create the worksheet
	book.create_worksheet :name => '客戶每日交易明細'

	#Create the rows to be inserted, and add row
	pnamelist=[]
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]+[newpname+'中']
	}
	row = ['類別']+pnamelist
	book.worksheet(0).insert_row(0, row)

	currentlist=[]
	winninglist=[]
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		#p cid
		#p pid
		#p swiftcode
		currentlist[pid=>1]
	}
	p currentlist
	#row = ['牌支']+currentlist
	#book.worksheet(0).insert_row(1, row)



	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	#Write the file
	year='2016'
	month='05'
	a='a'
	b='b'
	c='c'
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	book.write("report/#{year}/#{month}/#{a}_#{b}_#{c}_客戶每日交易明細.xls")

	#End
	p "#{a}_#{b}_#{c}_客戶每日交易明細.xls 已輸出"

end






@connection = WIN32OLE.new('ADODB.Connection')
@connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=main.mdb')



CustomDailyTransactionDetail(@connection,'100','1','2016/01/11')