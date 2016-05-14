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
	sql = "select * from [order] where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate='#{current_date}';"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select top 1 * from price where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate<='#{current_date}' order by CurrentDate desc;"
	recordset_price.Open(sql, connection)


	data_product = recordset_product.GetRows.transpose
	data_product.each(){|pid,pname|
		p pid,pname
	}
	recordset_product.close

=begin
	data_custom = recordset_custom.GetRows.transpose
	data_custom.each(){|cid,cname|
		p cid,cname
	}
	recordset_custom.close
=end


=begin
	data_order = recordset_order.GetRows.transpose
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group|
		p swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group
	}
	recordset_order.close
=end

=begin
	data_price = recordset_price.GetRows.transpose
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		p swiftcode,cid,pid,currentdate,currentprice,winningprice,upset

	}
	recordset_price.close
=end
end






@connection = WIN32OLE.new('ADODB.Connection')
@connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=main.mdb')



CustomDailyTransactionDetail(@connection,'100','1','2016/01/11')



=begin
# Begin Test
print "Spreadsheet Test\n"

# Create the rows to be inserted
row_1 = ['A1', 'B1']
row_2 = ['A2', 'B2']

# Create a new Workbook
new_book = Spreadsheet::Workbook.new

# Create the worksheet
new_book.create_worksheet :name => 'Sheet Name'

# Add row_1
new_book.worksheet(0).insert_row(0, row_1)

# Write the file
new_book.write('test.xls')

# End Test
print "Test Complete.\n"
=end


