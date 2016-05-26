# -*- coding: utf-8 -*-
require 'win32ole'  
#require 'spreadsheet' old function.
require 'writeexcel'




#客戶每日交易明細
def CustomDailyTransactionDetail(connection,current_pid,current_cid,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
 	if(current_pid.to_i()>=100)
    	sql = "select * from product where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} order by CLng(PID);"
	else
		sql = "select * from product where PID='#{current_pid}' order by CLng(PID);"
	end
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select top 1 * from custom where CID='#{current_cid}';"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from [order] where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CID='#{current_cid}' and CurrentDate='#{current_date}' order by group,CLng(PID);"
	else
		sql = "select * from [order] where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate='#{current_date}' order by group,CLng(PID);"
	end
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from price where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CID='#{current_cid}' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	else
		sql = "select * from price where PID='#{current_pid}' and CID='#{current_cid}' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	end
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	if(current_pid.to_i()>=100)
		full_pname=data_product.first.last.sub(/_.*/,'_全')
	else
		full_pname=data_product.first.last
	end
	pname=data_product.first.last.sub(/_.*/,'')


	#custom
	data_custom = recordset_custom.GetRows.transpose
	hash_custom=Hash.new
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
 		hash_custom["cid"]=cid
		hash_custom["cname"]=cname
		hash_custom["ctype"]=ctype
		hash_custom["address"]=address
		hash_custom["opendate"]=opendate
		hash_custom["bankid"]=bankid
		hash_custom["proportion"]=proportion
		hash_custom["bonustarget"]=bonustarget
		hash_custom["phone1"]=phone1
		hash_custom["phone2"]=phone2
		hash_custom["phone3"]=phone3
		hash_custom["phone4"]=phone4
		hash_custom["phone5"]=phone5
		hash_custom["phone6"]=phone6
		hash_custom["note"]=note
	}


	#price
	data_price = recordset_price.GetRows.transpose
	newcurrentdate=data_price.first.at(3)	#get newest currentdate from data array
	data_price.delete_if{|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		currentdate!=newcurrentdate
	}
	current_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,currentprice]
		}
	]
	winning_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,winningprice]
		}
	]

	#order
	data_order = recordset_order.GetRows.transpose


	
	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_#{hash_custom['cname']}_#{full_pname}_客戶每日交易明細.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Write a number and formula using A1 notation, and add row
	#列舉產品名
	pnamelist=Array.new
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]+[newpname+'中']
	}
	row = ['類別']+pnamelist
	sheet.write('A1', row)
	

	#列舉數量，並計算金額
	currentcountlist=Hash.new(0)
	winningcountlist=Hash.new(0)
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		currentcountlist[pid]+=currentcount.to_f()
		winningcountlist[pid]+=winningcount.to_f()
	}
	countlist=Array.new
	paylist=Array.new
	currentcountlist.each(){|key,value|
		countlist+=[currentcountlist[key]]+[winningcountlist[key]]
		paylist+=[currentcountlist[key]*current_price[key].to_f()]+[winningcountlist[key]*winning_price[key].to_f()]
	}
	row = ['牌支']+countlist	
	sheet.write('A2', row)
	row = ['金額']+paylist
	sheet.write('A3', row)


	#計算應收、漲價、佔成
	#row = ['應收']+[paylist.sum.round(4)]+['','','','']+['漲價']+[data_order.at(0).at(6)]+['']+['佔成']+[data_order.at(0).at(7)]	#get newest addmoney and bounsmoney from data array
	row = ['應收']
	sheet.write('A4',row)
	row = [paylist.inject(0){|sum,x| sum + x }.round(4)]	#inject value to sum
	sheet.merge_range('B4:F4',row,book.add_format)
	row= ['漲價']
	sheet.write('G4',row)
	row=[data_order.at(0).at(6)]
	sheet.merge_range('H4:I4',row,book.add_format)
	if(hash_custom["proportion"]==nil)
		row=['佔成']+['0%']
	else
		row=['佔成']+[hash_custom["proportion"].to_s()+'%']
	end
	sheet.write('J4',row)


	#顯示日期、產品類別、姓名
	row = [current_date]
	sheet.write('A5', row)
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	row = [pname]
	sheet.write('B5', row, format)
	row = [hash_custom['cname']]
	sheet.merge_range('C5:K5', row, book.add_format)

	#列舉產品名
	row = ['單號']+pnamelist
	sheet.write('A6', row)


	#為每一筆交易列舉
	currentcountlist=Hash.new(0)
	winningcountlist=Hash.new(0)
	oldgroup=-1
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if oldgroup==-1
			oldgroup=group
		elsif oldgroup!=group
			oldgroup=group

			countlist=Array.new
			currentcountlist.each(){|key,value|
				countlist+=[currentcountlist[key]]+[winningcountlist[key]]
			}
			row = ["%02d"%oldgroup]+countlist
			sheet.write('A'+(6+oldgroup).to_s(), row)
			
			currentcountlist=Hash.new(0)
			winningcountlist=Hash.new(0)
		end

		currentcountlist[pid]+=currentcount.to_f()
		winningcountlist[pid]+=winningcount.to_f()		
	}
	oldgroup=oldgroup.to_i()+1

	countlist=Array.new
	currentcountlist.each(){|key,value|
		countlist+=[currentcountlist[key]]+[winningcountlist[key]]
	}
	row = ["%02d"%oldgroup]+countlist	
	sheet.write('A'+(6+oldgroup).to_s(), row)



	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close

	#End
	p "#{free_current_date}_#{hash_custom['cname']}_#{full_pname}_客戶每日交易明細.xls 已輸出"

end






#每日交易加總表
def DailyTransactionCounting(connection,current_pid,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
 	if(current_pid.to_i()>=100)
    	sql = "select * from product where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} order by CLng(PID);"
	else
		sql = "select * from product where PID='#{current_pid}' order by CLng(PID);"
	end
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from [order] where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CurrentDate='#{current_date}' order by group,CLng(PID);"
	else
		sql = "select * from [order] where PID='#{current_pid}' and CurrentDate='#{current_date}' order by group,CLng(PID);"
	end
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from price where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	else
		sql = "select * from price where PID='#{current_pid}' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	end
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	if(current_pid.to_i()>=100)
		full_pname=data_product.first.last.sub(/_.*/,'_全')
	else
		full_pname=data_product.first.last
	end
	pname=data_product.first.last.sub(/_.*/,'')


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	newcurrentdate=data_price.first.at(3)	#get newest currentdate from data array
	data_price.delete_if{|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		currentdate!=newcurrentdate
	}
	current_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,currentprice]
		}
	]
	winning_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,winningprice]
		}
	]

	#order
	data_order = recordset_order.GetRows.transpose


	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_#{full_pname}_每日交易加總表.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#顯示標題
	row = [current_date]+[pname,'小計']
	sheet.write('A1', row)


	#列舉產品名
	pnamelist=Array.new
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]+[newpname+'中']
	}
	row = ['類別']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('A2', row, format)


	#計算出、入、留
	outcurrentcountlist=Hash.new(0)
	outwinningcountlist=Hash.new(0)
	incurrentcountlist=Hash.new(0)
	inwinningcountlist=Hash.new(0)
	customlist=Hash.new
	outaddmoney=0
	outbonusmoney=0
	inaddmoney=0
	inbonusmoney=0
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			outcurrentcountlist[pid]+=currentcount.to_f()
			outwinningcountlist[pid]+=winningcount.to_f()
			outaddmoney+=addmoney.to_f()
			outbonusmoney+=bonusmoney.to_f()
			
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['outaddmoney']+=addmoney.to_f()
			customlist[cid]['outbonusmoney']+=bonusmoney.to_f()
		else
			incurrentcountlist[pid]+=currentcount.to_f()
			inwinningcountlist[pid]+=winningcount.to_f()
			inaddmoney+=addmoney.to_f()
			inbonusmoney+=bonusmoney.to_f()

			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['inaddmoney']+=addmoney.to_f()
			customlist[cid]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney/=5
	outbonusmoney/=5
	inaddmoney/=5
	inbonusmoney/=5

	#計算誤差
	outpaylist=Array.new
	inpaylist=Array.new
	outcurrentcountlist.each(){|key,value|
		outpaylist+=[outcurrentcountlist[key]*current_price[key].to_f()]+[outwinningcountlist[key]*winning_price[key].to_f()]
		inpaylist+=[incurrentcountlist[key]*current_price[key].to_f()]+[inwinningcountlist[key]*winning_price[key].to_f()]
	}

	symbol=3
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"
	row = ['出']+outpaylist+[0,0,sumwithwater,sumwithoutwater,outaddmoney,outbonusmoney]
	sheet.write('A3', row)

	symbol=4
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"
	row = ['入']+inpaylist+[0,0,sumwithwater,sumwithoutwater,inaddmoney,inbonusmoney]
	sheet.write('A4', row)

	symbol=5
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"
	row = ['留']+[0,0,0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A5', row)

	row = ['誤差']
	(pnamelist.count+6).times(){|i|
		symbol=('B'.ord+i).chr
		row+=["=#{symbol}4-#{symbol}3-#{symbol}5"]
	}
	sheet.write('A6', row)

	#插入空白行
	row=['']
	sheet.write('A7', row)	

	#顯示出的客戶詳細清單
	row = ['出']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A8', row, format)

	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['outaddmoney']/=5
			customlist[cid]['outbonusmoney']/=5

			outcurrentcountlist.each(){|key,value|
				row+=[customlist[cid]['outcurrentcountlist'][key]*current_price[key].to_f()]+[customlist[cid]['outwinningcountlist'][key]*winning_price[key].to_f()]
			}
			row+=[0,0,sumwithwater,sumwithoutwater,customlist[cid]['outaddmoney'],customlist[cid]['outbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=9+rowindex
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"
	row = ['留底']+[0,0,0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A'+(9+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A'+(9+rowindex).to_s(), row, format)
	rowindex+=1

	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+P#{symbol}+Q#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+L#{symbol}-M#{symbol}+Q#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['inaddmoney']/=5
			customlist[cid]['inbonusmoney']/=5

			outcurrentcountlist.each(){|key,value|
				row+=[customlist[cid]['incurrentcountlist'][key]*current_price[key].to_f()]+[customlist[cid]['inwinningcountlist'][key]*winning_price[key].to_f()]
			}
			row+=[0,0,sumwithwater,sumwithoutwater,customlist[cid]['inaddmoney'],customlist[cid]['inbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_#{full_pname}_每日交易加總表.xls 已輸出"

end





#全產品每日交易加總表
def AllDailyTransactionCounting(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where PID not like '%5' order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from [order] where PID not like '%5' and CurrentDate='#{current_date}' order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where CLng(PID)<100 and PID not like '%5' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='日總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	newcurrentdate=data_price.first.at(3)	#get newest currentdate from data array
	data_price.delete_if{|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		currentdate!=newcurrentdate
	}
	current_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,currentprice]
		}
	]
	winning_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,winningprice]
		}
	]

	#order
	data_order = recordset_order.GetRows.transpose

	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_全產品每日日總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#顯示標題
	row = [current_date]+['日總計']
	sheet.write('A1', row)


	#列舉產品名
	pnamelist=Array.new
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]+[newpname+'中'] if(!pnamelist.include?(newpname))
	}
	row = ['類別']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('A2', row, format)


	#計算出、入、留
	outcurrentcountlist=Hash.new(0)
	outwinningcountlist=Hash.new(0)
	incurrentcountlist=Hash.new(0)
	inwinningcountlist=Hash.new(0)
	customlist=Hash.new
	outaddmoney=0
	outbonusmoney=0
	inaddmoney=0
	inbonusmoney=0
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			outcurrentcountlist[pid]+=currentcount.to_f()
			outwinningcountlist[pid]+=winningcount.to_f()
			outaddmoney+=addmoney.to_f()
			outbonusmoney+=bonusmoney.to_f()
			
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['outaddmoney']+=addmoney.to_f()
			customlist[cid]['outbonusmoney']+=bonusmoney.to_f()
		else
			incurrentcountlist[pid]+=currentcount.to_f()
			inwinningcountlist[pid]+=winningcount.to_f()
			inaddmoney+=addmoney.to_f()
			inbonusmoney+=bonusmoney.to_f()

			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['inaddmoney']+=addmoney.to_f()
			customlist[cid]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney/=4
	outbonusmoney/=4
	inaddmoney/=4
	inbonusmoney/=4

	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	outgetlist=Hash.new(0)
	ingetlist=Hash.new(0)
	outcurrentcountlist.each(){|key,value|
		outpaylist[key[-1]]+=outcurrentcountlist[key]*current_price[key].to_f()
		outgetlist[key[-1]]+=outwinningcountlist[key]*winning_price[key].to_f()
		inpaylist[key[-1]]+=incurrentcountlist[key]*current_price[key].to_f()
		ingetlist[key[-1]]+=inwinningcountlist[key]*winning_price[key].to_f()
	}
	outlist=[outpaylist['1'],outgetlist['1'],outpaylist['2'],outgetlist['2'],outpaylist['3'],outgetlist['3'],outpaylist['4'],outgetlist['4']]
	inlist=[inpaylist['1'],ingetlist['1'],inpaylist['2'],ingetlist['2'],inpaylist['3'],ingetlist['3'],inpaylist['4'],ingetlist['4']]

	symbol=3
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"
	row = ['出']+outlist+[0,0,sumwithwater,sumwithoutwater,outaddmoney,outbonusmoney]
	sheet.write('A3', row)

	symbol=4
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"
	row = ['入']+inlist+[0,0,sumwithwater,sumwithoutwater,inaddmoney,inbonusmoney]
	sheet.write('A4', row)

	symbol=5
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"
	row = ['留']+[0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A5', row)

	row = ['誤差']
	(pnamelist.count+6).times(){|i|
		symbol=('B'.ord+i).chr
		row+=["=#{symbol}4-#{symbol}3-#{symbol}5"]
	}
	sheet.write('A6', row)

	#插入空白行
	row=['']
	sheet.write('A7', row)	

	#顯示出的客戶詳細清單
	row = ['出']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A8', row, format)

	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['outaddmoney']/=4
			customlist[cid]['outbonusmoney']/=4


			outpaylist=Hash.new(0)
			outgetlist=Hash.new(0)
			outcurrentcountlist.each(){|key,value|
				outpaylist[key[-1]]+=outcurrentcountlist[key]*current_price[key].to_f()
				outgetlist[key[-1]]+=outwinningcountlist[key]*winning_price[key].to_f()
			}
			outlist=[outpaylist['1'],outgetlist['1'],outpaylist['2'],outgetlist['2'],outpaylist['3'],outgetlist['3'],outpaylist['4'],outgetlist['4']]			
			row = ['出']+outlist+[0,0,sumwithwater,sumwithoutwater,customlist[cid]['outaddmoney'],customlist[cid]['outbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=9+rowindex
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"
	row = ['留底']+[0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A'+(9+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A'+(9+rowindex).to_s(), row, format)
	rowindex+=1

	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+N#{symbol}+O#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+F#{symbol}-G#{symbol}+H#{symbol}-I#{symbol}+J#{symbol}-K#{symbol}+O#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['inaddmoney']/=4
			customlist[cid]['inbonusmoney']/=4


			inpaylist=Hash.new(0)
			ingetlist=Hash.new(0)
			incurrentcountlist.each(){|key,value|
				inpaylist[key[-1]]+=incurrentcountlist[key]*current_price[key].to_f()
				ingetlist[key[-1]]+=inwinningcountlist[key]*winning_price[key].to_f()
			}
			inlist=[inpaylist['1'],ingetlist['1'],inpaylist['2'],ingetlist['2'],inpaylist['3'],ingetlist['3'],inpaylist['4'],ingetlist['4']]
			row = ['入']+inlist+[0,0,sumwithwater,sumwithoutwater,customlist[cid]['inaddmoney'],customlist[cid]['inbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_全產品每日日總計.xls 已輸出"

end








#全產品一週交易加總表
def AllWeekTransactionCounting(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where CLng(PID)<100 order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)
	
	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	begin_date=(Date.parse(current_date)-6).to_s().gsub('-','/')
	sql = "select * from [order] where (CurrentDate>='#{begin_date}' and CurrentDate<='#{current_date}') order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='週總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	current_price=Hash.new()
	winning_price=Hash.new()
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		current_price[currentdate]=Hash.new() if current_price[currentdate]==nil
		current_price[currentdate][pid]=currentprice

		winning_price[currentdate]=Hash.new() if winning_price[currentdate]==nil
		winning_price[currentdate][pid]=winningprice
	}
	#倒序依日期排序
	current_price.sort_by{|key,v| key}.reverse
	winning_price.sort_by{|key,v| key}.reverse

	#order
	data_order = recordset_order.GetRows.transpose


	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_全產品一週週總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#列舉日期
	row = [begin_date+'~'+current_date]
	sheet.write('A1', row)
	row = ['D1','D2','D3','D4','D5','D6','D7']
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('B1', row, format)
	row = ['總計','應收','前帳','水']
	sheet.write('I1', row)


	#計算出、入、留
	outcurrentcountlist=Hash.new()
	outwinningcountlist=Hash.new()
	incurrentcountlist=Hash.new()
	inwinningcountlist=Hash.new()
	customlist=Hash.new
	outaddmoney=Hash.new(0)
	outbonusmoney=Hash.new(0)
	inaddmoney=Hash.new(0)
	inbonusmoney=Hash.new(0)
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			outcurrentcountlist[currentdate][pid]+=currentcount.to_f()
			outwinningcountlist[currentdate][pid]+=winningcount.to_f()
			outaddmoney[current_date]+=addmoney.to_f()
			outbonusmoney[current_date]+=bonusmoney.to_f()
			
			
			customlist[cid][current_date]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['outaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['outbonusmoney']+=bonusmoney.to_f()
		else
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			incurrentcountlist[currentdate][pid]+=currentcount.to_f()
			inwinningcountlist[currentdate][pid]+=winningcount.to_f()
			inaddmoney[current_date]+=addmoney.to_f()
			inbonusmoney[current_date]+=bonusmoney.to_f()


			customlist[cid][current_date]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['inaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney.each(){|key,value|
		outaddmoney[key]/=5
		outbonusmoney[key]/=5
		inaddmoney[key]/=5
		inbonusmoney[key]/=5	
	}


	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	(begin_date..current_date).each(){|key|
		#找出過出日期相對最接近且有輸入價格的
		mostneardate=nil
		current_price.each(){|key1,phash1|
			if(Date.parse(key1)<=Date.parse(key))
				mostneardate=key1
				break
			end
		}


		#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
		if(mostneardate!=nil)
			if(outcurrentcountlist[key]!=nil)
				data_product.each(){|pid,pname|
					outpaylist[key]+=(outcurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()-outwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
					inpaylist[key]+=(incurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()--inwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
				}
			else
				outpaylist[key]=0
				inpaylist[key]=0
			end
		else
			outpaylist[key]=0
			inpaylist[key]=0
		end	
	}

	symbol=2
	sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
	sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"	
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['出']+outpaylist.collect(){|key,value| value}+[sumwitoutholdpay,sumwitholdpay,0,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]	
	sheet.write('A2', row)

	symbol=3
	sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
	sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"	
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['入']+inpaylist.collect(){|key,value| value}+[sumwitoutholdpay,sumwitholdpay,0,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]	
	sheet.write('A3', row)

	symbol=4
	sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
	sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"	
	row = ['留']+[0,0,0,0,0,0,0,sumwitoutholdpay,sumwitholdpay,0,0]
	sheet.write('A4', row)

	row = ['誤差']
	(11).times(){|i|
		symbol=('B'.ord+i).chr
		row+=["=#{symbol}3-#{symbol}2-#{symbol}4"]
	}
	sheet.write('A5', row)

	#插入空白行
	row=['']
	sheet.write('A6', row)	

	#顯示出的客戶詳細清單
	row = ['出']
	sheet.write('A7', row)
	row = ['D1','D2','D3','D4','D5','D6','D7']
	sheet.write('B7', row, format)
	row = ['總計','應收','前帳','水']
	sheet.write('I7', row)

	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
		sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"	

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,sumwitoutholdpay,sumwitholdpay,0,0]
		else
			outpaylist=Hash.new(0)
			(begin_date..current_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(outcurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							outpaylist[key]+=(customlist[cid][current_date]['outcurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['outwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						outpaylist[key]=0
					end
				else
					outpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			outaddmoney.each(){|key,value|
				customlist[cid][current_date]['outaddmoney']/=5
				customlist[cid][current_date]['outbonusmoney']/=5
			}
		

			row += outpaylist.collect(){|key,value| value}+[sumwitoutholdpay,sumwitholdpay,0,customlist[cid][current_date]['outaddmoney']]	
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=8+rowindex
	sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
	sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"
	row = ['留底']+[0,0,0,0,0,0,0,sumwitoutholdpay,sumwitholdpay,0,0]
	sheet.write('A'+(8+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']
	sheet.write('A'+(8+rowindex).to_s(), row)
	row = ['D1','D2','D3','D4','D5','D6','D7']
	sheet.write('B'+(8+rowindex).to_s(), row, format)
	row = ['總計','應收','前帳','水']
	sheet.write('I'+(8+rowindex).to_s(), row)
	rowindex+=1

	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutholdpay="=SUM(B#{symbol}:H#{symbol})"
		sumwitholdpay="=SUM(B#{symbol}:H#{symbol})+K#{symbol}+L#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,0,0,0,sumwitoutholdpay,sumwitholdpay,0,0]
		else
			inpaylist=Hash.new(0)
			(begin_date..current_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(incurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							inpaylist[key]+=(customlist[cid][current_date]['incurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['inwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						inpaylist[key]=0
					end
				else
					inpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			inaddmoney.each(){|key,value|
				customlist[cid][current_date]['inaddmoney']/=5
				customlist[cid][current_date]['inbonusmoney']/=5
			}
		

			row += inpaylist.collect(){|key,value| value}+[sumwitoutholdpay,sumwitholdpay,0,customlist[cid][current_date]['inaddmoney']]	
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_全產品一週週總計.xls 已輸出"

end








#全產品當月交易加總表
def AllMonthTransactionCounting(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')



	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where CLng(PID)<100 order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)
	
	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	begin_date=Date.civil(year.to_i(), month.to_i(), 1).to_s().gsub('-','/')
	end_date=Date.civil(year.to_i(), month.to_i(), -1).to_s().gsub('-','/')
	sql = "select * from [order] where (CurrentDate>='#{begin_date}' and CurrentDate<='#{current_date}') order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='月總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	current_price=Hash.new()
	winning_price=Hash.new()
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		current_price[currentdate]=Hash.new() if current_price[currentdate]==nil
		current_price[currentdate][pid]=currentprice

		winning_price[currentdate]=Hash.new() if winning_price[currentdate]==nil
		winning_price[currentdate][pid]=winningprice
	}
	#倒序依日期排序
	current_price.sort_by{|key,v| key}.reverse
	winning_price.sort_by{|key,v| key}.reverse

	#order
	data_order = recordset_order.GetRows.transpose



	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_全產品當月月總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#列舉日期
	maxdate=0
	symbol='A'
	row = [begin_date+'~'+current_date]
	sheet.write('A1', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
		maxdate=index
	}
	maxdate+=1
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('B1', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}1", row)


	#計算出、入、留
	outcurrentcountlist=Hash.new()
	outwinningcountlist=Hash.new()
	incurrentcountlist=Hash.new()
	inwinningcountlist=Hash.new()
	customlist=Hash.new
	outaddmoney=Hash.new(0)
	outbonusmoney=Hash.new(0)
	inaddmoney=Hash.new(0)
	inbonusmoney=Hash.new(0)
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			outcurrentcountlist[currentdate][pid]+=currentcount.to_f()
			outwinningcountlist[currentdate][pid]+=winningcount.to_f()
			outaddmoney[current_date]+=addmoney.to_f()
			outbonusmoney[current_date]+=bonusmoney.to_f()
			
			
			customlist[cid][current_date]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['outaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['outbonusmoney']+=bonusmoney.to_f()
		else
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			incurrentcountlist[currentdate][pid]+=currentcount.to_f()
			inwinningcountlist[currentdate][pid]+=winningcount.to_f()
			inaddmoney[current_date]+=addmoney.to_f()
			inbonusmoney[current_date]+=bonusmoney.to_f()


			customlist[cid][current_date]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['inaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney.each(){|key,value|
		outaddmoney[key]/=5
		outbonusmoney[key]/=5
		inaddmoney[key]/=5
		inbonusmoney[key]/=5	
	}


	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	(begin_date..end_date).each(){|key|
		#找出過出日期相對最接近且有輸入價格的
		mostneardate=nil
		current_price.each(){|key1,phash1|
			if(Date.parse(key1)<=Date.parse(key))
				mostneardate=key1
				break
			end
		}


		#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
		if(mostneardate!=nil)
			if(outcurrentcountlist[key]!=nil)
				data_product.each(){|pid,pname|
					outpaylist[key]+=(outcurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()-outwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
					inpaylist[key]+=(incurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()--inwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
				}
			else
				outpaylist[key]=0
				inpaylist[key]=0
			end
		else
			outpaylist[key]=0
			inpaylist[key]=0
		end	
	}

	symbol=2
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['出']+outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A2', row)

	symbol=3
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['入']+inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A3', row)

	symbol=4
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A4', row)

	row = ['誤差']
	('B'..'Z').each(){|i|
		row+=["=#{i}3-#{i}2-#{i}4"]
	}
	row+=["=AA3-AA2-AA4"]+["=AB3-AB2-AB4"]+["=AC3-AC2-AC4"]+["=AD3-AD2-AD4"]+["=AE3-AE2-AE4"]+["=AF3-AF2-AF4"]+["=AG3-AG2-AG4"]
	row+=["=AH3-AH2-AH4"] if(maxdate>29)
	row+=["=AI3-AI2-AI4"] if(maxdate>30)
	sheet.write('A5', row)

	#插入空白行
	row=['']
	sheet.write('A6', row)	

	#顯示出的客戶詳細清單
	row = ['出']
	sheet.write('A7', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B7', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}7", row)	


	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			outpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(outcurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							outpaylist[key]+=(customlist[cid][current_date]['outcurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['outwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						outpaylist[key]=0
					end
				else
					outpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			outaddmoney.each(){|key,value|
				customlist[cid][current_date]['outaddmoney']/=5
				customlist[cid][current_date]['outbonusmoney']/=5
			}
		

			row += outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=8+rowindex
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留底']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,0,0]
	sheet.write('A'+(8+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']
	sheet.write('A'+(8+rowindex).to_s(), row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B'+(8+rowindex).to_s(), row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}#{(8+rowindex)}", row)	
	rowindex+=1


	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			inpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(incurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							inpaylist[key]+=(customlist[cid][current_date]['incurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['inwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						inpaylist[key]=0
					end
				else
					inpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			inaddmoney.each(){|key,value|
				customlist[cid][current_date]['inaddmoney']/=5
				customlist[cid][current_date]['inbonusmoney']/=5
			}
		

			row += inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_全產品當月月總計.xls 已輸出"

end






#當月交易加總表
def MonthTransactionCounting(connection,current_pid,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')



	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
    	sql = "select * from product where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} order by CLng(PID);"
	else
		sql = "select * from product order by CLng(PID);"
	end
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)
	
	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	begin_date=Date.civil(year.to_i(), month.to_i(), 1).to_s().gsub('-','/')
	end_date=Date.civil(year.to_i(), month.to_i(), -1).to_s().gsub('-','/')
	if(current_pid.to_i()>=100)
		sql = "select * from [order] where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and (CurrentDate>='#{begin_date}' and CurrentDate<='#{current_date}') order by group,CLng(PID);"
	else
		sql = "select * from [order] where (CurrentDate>='#{begin_date}' and CurrentDate<='#{current_date}') order by group,CLng(PID);"
	end
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	if(current_pid.to_i()>=100)
		sql = "select * from price where cint(PID)>=#{current_pid[1]+'0'} and cint(PID)<=#{current_pid[1]+'9'} and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	else
		sql = "select * from price where CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	end
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	if(current_pid.to_i()>=100)
		full_pname=data_product.first.last.sub(/_.*/,'_全')
	else
		full_pname=data_product.first.last
	end
	pname='月總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	current_price=Hash.new()
	winning_price=Hash.new()
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		current_price[currentdate]=Hash.new() if current_price[currentdate]==nil
		current_price[currentdate][pid]=currentprice

		winning_price[currentdate]=Hash.new() if winning_price[currentdate]==nil
		winning_price[currentdate][pid]=winningprice
	}
	#倒序依日期排序
	current_price.sort_by{|key,v| key}.reverse
	winning_price.sort_by{|key,v| key}.reverse

	#order
	data_order = recordset_order.GetRows.transpose



	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_#{full_pname}_當月月總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#列舉日期
	maxdate=0
	symbol='A'
	row = [begin_date+'~'+current_date]
	sheet.write('A1', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
		maxdate=index
	}
	maxdate+=1
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('B1', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}1", row)


	#計算出、入、留
	outcurrentcountlist=Hash.new()
	outwinningcountlist=Hash.new()
	incurrentcountlist=Hash.new()
	inwinningcountlist=Hash.new()
	customlist=Hash.new
	outaddmoney=Hash.new(0)
	outbonusmoney=Hash.new(0)
	inaddmoney=Hash.new(0)
	inbonusmoney=Hash.new(0)
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			outcurrentcountlist[currentdate][pid]+=currentcount.to_f()
			outwinningcountlist[currentdate][pid]+=winningcount.to_f()
			outaddmoney[current_date]+=addmoney.to_f()
			outbonusmoney[current_date]+=bonusmoney.to_f()
			
			
			customlist[cid][current_date]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['outaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['outbonusmoney']+=bonusmoney.to_f()
		else
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			incurrentcountlist[currentdate][pid]+=currentcount.to_f()
			inwinningcountlist[currentdate][pid]+=winningcount.to_f()
			inaddmoney[current_date]+=addmoney.to_f()
			inbonusmoney[current_date]+=bonusmoney.to_f()


			customlist[cid][current_date]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['inaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney.each(){|key,value|
		outaddmoney[key]/=5
		outbonusmoney[key]/=5
		inaddmoney[key]/=5
		inbonusmoney[key]/=5	
	}


	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	(begin_date..end_date).each(){|key|
		#找出過出日期相對最接近且有輸入價格的
		mostneardate=nil
		current_price.each(){|key1,phash1|
			if(Date.parse(key1)<=Date.parse(key))
				mostneardate=key1
				break
			end
		}


		#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
		if(mostneardate!=nil)
			if(outcurrentcountlist[key]!=nil)
				data_product.each(){|pid,pname|
					outpaylist[key]+=(outcurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()-outwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
					inpaylist[key]+=(incurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()--inwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
				}
			else
				outpaylist[key]=0
				inpaylist[key]=0
			end
		else
			outpaylist[key]=0
			inpaylist[key]=0
		end	
	}

	symbol=2
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['出']+outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A2', row)

	symbol=3
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['入']+inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A3', row)

	symbol=4
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A4', row)

	row = ['誤差']
	('B'..'Z').each(){|i|
		row+=["=#{i}3-#{i}2-#{i}4"]
	}
	row+=["=AA3-AA2-AA4"]+["=AB3-AB2-AB4"]+["=AC3-AC2-AC4"]+["=AD3-AD2-AD4"]+["=AE3-AE2-AE4"]+["=AF3-AF2-AF4"]+["=AG3-AG2-AG4"]
	row+=["=AH3-AH2-AH4"] if(maxdate>29)
	row+=["=AI3-AI2-AI4"] if(maxdate>30)
	sheet.write('A5', row)

	#插入空白行
	row=['']
	sheet.write('A6', row)	

	#顯示出的客戶詳細清單
	row = ['出']
	sheet.write('A7', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B7', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}7", row)	


	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			outpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(outcurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							outpaylist[key]+=(customlist[cid][current_date]['outcurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['outwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						outpaylist[key]=0
					end
				else
					outpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			outaddmoney.each(){|key,value|
				customlist[cid][current_date]['outaddmoney']/=5
				customlist[cid][current_date]['outbonusmoney']/=5
			}
		

			row += outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=8+rowindex
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留底']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,0,0]
	sheet.write('A'+(8+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']
	sheet.write('A'+(8+rowindex).to_s(), row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B'+(8+rowindex).to_s(), row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}#{(8+rowindex)}", row)	
	rowindex+=1


	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			inpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(incurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							inpaylist[key]+=(customlist[cid][current_date]['incurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['inwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						inpaylist[key]=0
					end
				else
					inpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			inaddmoney.each(){|key,value|
				customlist[cid][current_date]['inaddmoney']/=5
				customlist[cid][current_date]['inbonusmoney']/=5
			}
		

			row += inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_#{full_pname}_當月月總計.xls 已輸出"

end











#全產品4K每日交易加總表
def AllDaily4KTransactionCounting(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where PID like '%4' order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from [order] where PID like '%4' and CurrentDate='#{current_date}' order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where CLng(PID)<100 and PID like '%4' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='日總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	newcurrentdate=data_price.first.at(3)	#get newest currentdate from data array
	data_price.delete_if{|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		currentdate!=newcurrentdate
	}
	current_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,currentprice]
		}
	]
	winning_price=Hash[
		data_price.collect(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
			[pid,winningprice]
		}
	]

	#order
	data_order = recordset_order.GetRows.transpose

	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_全產品4K每日日總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#顯示標題
	row = [current_date]+['4K日總計']
	sheet.write('A1', row)


	#列舉產品名
	pnamelist=Array.new
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]+[newpname+'中'] if(!pnamelist.include?(newpname))
	}
	row = ['類別']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('A2', row, format)


	#計算出、入、留
	outcurrentcountlist=Hash.new(0)
	outwinningcountlist=Hash.new(0)
	incurrentcountlist=Hash.new(0)
	inwinningcountlist=Hash.new(0)
	customlist=Hash.new
	outaddmoney=0
	outbonusmoney=0
	inaddmoney=0
	inbonusmoney=0
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			outcurrentcountlist[pid]+=currentcount.to_f()
			outwinningcountlist[pid]+=winningcount.to_f()
			outaddmoney+=addmoney.to_f()
			outbonusmoney+=bonusmoney.to_f()
			
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['outaddmoney']+=addmoney.to_f()
			customlist[cid]['outbonusmoney']+=bonusmoney.to_f()
		else
			incurrentcountlist[pid]+=currentcount.to_f()
			inwinningcountlist[pid]+=winningcount.to_f()
			inaddmoney+=addmoney.to_f()
			inbonusmoney+=bonusmoney.to_f()

			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end
			customlist[cid]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid]['inaddmoney']+=addmoney.to_f()
			customlist[cid]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney/=4
	outbonusmoney/=4
	inaddmoney/=4
	inbonusmoney/=4

	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	outgetlist=Hash.new(0)
	ingetlist=Hash.new(0)
	outcurrentcountlist.each(){|key,value|
		outpaylist[key[-1]]+=outcurrentcountlist[key]*current_price[key].to_f()
		outgetlist[key[-1]]+=outwinningcountlist[key]*winning_price[key].to_f()
		inpaylist[key[-1]]+=incurrentcountlist[key]*current_price[key].to_f()
		ingetlist[key[-1]]+=inwinningcountlist[key]*winning_price[key].to_f()
	}
	outlist=[outpaylist['4'],outgetlist['4']]
	inlist=[inpaylist['4'],ingetlist['4']]

	symbol=3
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"
	row = ['出']+outlist+[0,0,sumwithwater,sumwithoutwater,outaddmoney,outbonusmoney]
	sheet.write('A3', row)

	symbol=4
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"
	row = ['入']+inlist+[0,0,sumwithwater,sumwithoutwater,inaddmoney,inbonusmoney]
	sheet.write('A4', row)

	symbol=5
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"
	row = ['留']+[0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A5', row)

	row = ['誤差']
	(pnamelist.count+6).times(){|i|
		symbol=('B'.ord+i).chr
		row+=["=#{symbol}4-#{symbol}3-#{symbol}5"]
	}
	sheet.write('A6', row)

	#插入空白行
	row=['']
	sheet.write('A7', row)	

	#顯示出的客戶詳細清單
	row = ['出']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A8', row, format)

	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['outaddmoney']/=4
			customlist[cid]['outbonusmoney']/=4


			outpaylist=Hash.new(0)
			outgetlist=Hash.new(0)
			outcurrentcountlist.each(){|key,value|
				outpaylist[key[-1]]+=outcurrentcountlist[key]*current_price[key].to_f()
				outgetlist[key[-1]]+=outwinningcountlist[key]*winning_price[key].to_f()
			}
			outlist=[outpaylist['4'],outgetlist['4']]			
			row = ['出']+outlist+[0,0,sumwithwater,sumwithoutwater,customlist[cid]['outaddmoney'],customlist[cid]['outbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=9+rowindex
	sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
	sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"
	row = ['留底']+[0,0,0,0,sumwithwater,sumwithoutwater,0,0]
	sheet.write('A'+(9+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']+pnamelist+['其它','其它中','應收(含水)','應收(扣水)','水','漲價']
	sheet.write('A'+(9+rowindex).to_s(), row, format)
	rowindex+=1

	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=9+rowindex
		sumwithwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+H#{symbol}+I#{symbol}"
		sumwithoutwater="=B#{symbol}-C#{symbol}+D#{symbol}-E#{symbol}+I#{symbol}"

		row = [cname]
		if(customlist[cid]==nil)
			row+=[0,0,0,0,sumwithwater,sumwithoutwater,0,0]
		else
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			customlist[cid]['inaddmoney']/=4
			customlist[cid]['inbonusmoney']/=4


			inpaylist=Hash.new(0)
			ingetlist=Hash.new(0)
			incurrentcountlist.each(){|key,value|
				inpaylist[key[-1]]+=incurrentcountlist[key]*current_price[key].to_f()
				ingetlist[key[-1]]+=inwinningcountlist[key]*winning_price[key].to_f()
			}
			inlist=[inpaylist['4'],ingetlist['4']]
			row = ['入']+inlist+[0,0,sumwithwater,sumwithoutwater,customlist[cid]['inaddmoney'],customlist[cid]['inbonusmoney']]
		end
		sheet.write('A'+(9+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_全產品4K每日日總計.xls 已輸出"

end





#全產品4K當月交易加總表
def AllMonth4KTransactionCounting(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')



	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where PID like '%4' order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)
	
	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	begin_date=Date.civil(year.to_i(), month.to_i(), 1).to_s().gsub('-','/')
	end_date=Date.civil(year.to_i(), month.to_i(), -1).to_s().gsub('-','/')
	sql = "select * from [order] where PID like '%4' and (CurrentDate>='#{begin_date}' and CurrentDate<='#{current_date}') order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where PID like '%4' and CurrentDate<='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='4K月總計'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	current_price=Hash.new()
	winning_price=Hash.new()
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		current_price[currentdate]=Hash.new() if current_price[currentdate]==nil
		current_price[currentdate][pid]=currentprice

		winning_price[currentdate]=Hash.new() if winning_price[currentdate]==nil
		winning_price[currentdate][pid]=winningprice
	}
	#倒序依日期排序
	current_price.sort_by{|key,v| key}.reverse
	winning_price.sort_by{|key,v| key}.reverse

	#order
	data_order = recordset_order.GetRows.transpose



	

	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_全產品4K當月月總計.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Create the rows to be inserted, and add row
	#列舉日期
	maxdate=0
	symbol='A'
	row = [begin_date+'~'+current_date]
	sheet.write('A1', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
		maxdate=index
	}
	maxdate+=1
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('B1', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}1", row)


	#計算出、入、留
	outcurrentcountlist=Hash.new()
	outwinningcountlist=Hash.new()
	incurrentcountlist=Hash.new()
	inwinningcountlist=Hash.new()
	customlist=Hash.new
	outaddmoney=Hash.new(0)
	outbonusmoney=Hash.new(0)
	inaddmoney=Hash.new(0)
	inbonusmoney=Hash.new(0)
	data_order.each(){|swiftcode,cid,pid,currentdate,currentcount,winningcount,addmoney,bonusmoney,note,group| 
		if(currentcount.to_f()>=0)
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			outcurrentcountlist[currentdate][pid]+=currentcount.to_f()
			outwinningcountlist[currentdate][pid]+=winningcount.to_f()
			outaddmoney[current_date]+=addmoney.to_f()
			outbonusmoney[current_date]+=bonusmoney.to_f()
			
			
			customlist[cid][current_date]['outcurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['outwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['outaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['outbonusmoney']+=bonusmoney.to_f()
		else
			#若計價清單要存的日期的HASH尚未建立，則為每個要存的日期建置HASH，以供別存到每一天，後續才能跟不同的價格表相乘
			if(outcurrentcountlist[currentdate]==nil)
				outcurrentcountlist[currentdate]=Hash.new(0)
				outwinningcountlist[currentdate]=Hash.new(0)
				incurrentcountlist[currentdate]=Hash.new(0)
				inwinningcountlist[currentdate]=Hash.new(0)
			end
			#若客戶的ID HASH尚未建立，則為每個客戶ID建置HASH，以供分別儲存他們的交易及中獎數量
			if(customlist[cid]==nil)
				customlist[cid]=Hash.new()
			end
			if(customlist[cid][currentdate]==nil)
				customlist[cid][currentdate]={'outcurrentcountlist'=>Hash.new(0),'outwinningcountlist'=>Hash.new(0),'outaddmoney'=>0,'outbonusmoney'=>0,'incurrentcountlist'=>Hash.new(0),'inwinningcountlist'=>Hash.new(0),'inaddmoney'=>0,'inbonusmoney'=>0}
			end


			incurrentcountlist[currentdate][pid]+=currentcount.to_f()
			inwinningcountlist[currentdate][pid]+=winningcount.to_f()
			inaddmoney[current_date]+=addmoney.to_f()
			inbonusmoney[current_date]+=bonusmoney.to_f()


			customlist[cid][current_date]['incurrentcountlist'][pid]+=currentcount.to_f()
			customlist[cid][current_date]['inwinningcountlist'][pid]+=winningcount.to_f()
			customlist[cid][current_date]['inaddmoney']+=addmoney.to_f()
			customlist[cid][current_date]['inbonusmoney']+=bonusmoney.to_f()
		end
	}
	#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
	outaddmoney.each(){|key,value|
		outaddmoney[key]/=5
		outbonusmoney[key]/=5
		inaddmoney[key]/=5
		inbonusmoney[key]/=5	
	}


	#計算誤差
	outpaylist=Hash.new(0)
	inpaylist=Hash.new(0)
	(begin_date..end_date).each(){|key|
		#找出過出日期相對最接近且有輸入價格的
		mostneardate=nil
		current_price.each(){|key1,phash1|
			if(Date.parse(key1)<=Date.parse(key))
				mostneardate=key1
				break
			end
		}


		#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
		if(mostneardate!=nil)
			if(outcurrentcountlist[key]!=nil)
				data_product.each(){|pid,pname|
					outpaylist[key]+=(outcurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()-outwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
					inpaylist[key]+=(incurrentcountlist[key][pid]*current_price[mostneardate][pid].to_f()--inwinningcountlist[key][pid]*winning_price[mostneardate][pid].to_f())
				}
			else
				outpaylist[key]=0
				inpaylist[key]=0
			end
		else
			outpaylist[key]=0
			inpaylist[key]=0
		end	
	}

	symbol=2
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['出']+outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A2', row)

	symbol=3
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	#PAYLIST是將HASH後面的數值轉為ARRAY，再COLLECT分散輸出成ARRAY
	#ADDMONEY是做同上的動作後，再把輸出的ARRAY透過INJECT的方式加總為一個數字
	row = ['入']+inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A3', row)

	symbol=4
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
	sheet.write('A4', row)

	row = ['誤差']
	('B'..'Z').each(){|i|
		row+=["=#{i}3-#{i}2-#{i}4"]
	}
	row+=["=AA3-AA2-AA4"]+["=AB3-AB2-AB4"]+["=AC3-AC2-AC4"]+["=AD3-AD2-AD4"]+["=AE3-AE2-AE4"]+["=AF3-AF2-AF4"]+["=AG3-AG2-AG4"]
	row+=["=AH3-AH2-AH4"] if(maxdate>29)
	row+=["=AI3-AI2-AI4"] if(maxdate>30)
	sheet.write('A5', row)

	#插入空白行
	row=['']
	sheet.write('A6', row)	

	#顯示出的客戶詳細清單
	row = ['出']
	sheet.write('A7', row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B7', row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}7", row)	


	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			outpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(outcurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							outpaylist[key]+=(customlist[cid][current_date]['outcurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['outwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						outpaylist[key]=0
					end
				else
					outpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			outaddmoney.each(){|key,value|
				customlist[cid][current_date]['outaddmoney']/=5
				customlist[cid][current_date]['outbonusmoney']/=5
			}
		

			row += outpaylist.collect(){|key,value| value}+[sumwitoutother,outaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},outbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}
	

	#顯示留底的客戶詳細清單
	symbol=8+rowindex
	sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"
	row = ['留底']
	maxdate.times(){
		row+=[0]
	}
	row+=[sumwitoutother,0,0]
	sheet.write('A'+(8+rowindex).to_s(), row)
	rowindex+=1


	#顯示入的客戶詳細清單
	row = ['入']
	sheet.write('A'+(8+rowindex).to_s(), row)
	row = []
	(begin_date..end_date).each_with_index(){|value,index|
		row+=['D'+(index+1).to_s()]
	}
	sheet.write('B'+(8+rowindex).to_s(), row, format)
	row = ['總計','水','漲價']
	if(maxdate==29)
		maxsymbol='AD'
		symbol='AE'
	elsif(maxdate==30)
		maxsymbol='AE'
		symbol='AF'
	elsif(maxdate==31)
		maxsymbol='AF'
		symbol='AG'
	end
	sheet.write("#{symbol}#{(8+rowindex)}", row)	
	rowindex+=1


	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=8+rowindex
		sumwitoutother="=SUM(B#{symbol}:#{maxsymbol}#{symbol})"

		row = [cname]
		if(customlist[cid]==nil)
			#依這個月有幾天，來填上許許多多的零
			maxdate.times(){
				row+=[0]
			}
			row+=[sumwitoutother,0,0]
		else
			inpaylist=Hash.new(0)
			(begin_date..end_date).each(){|key|
				#找出過出日期相對最接近且有輸入價格的
				mostneardate=nil
				current_price.each(){|key1,phash1|
					if(Date.parse(key1)<=Date.parse(key))
						mostneardate=key1
						break
					end
				}


				#列舉並加總該日所有的產品，如果沒有找到最接近日期的價格表，那就乾脆不加了
				if(mostneardate!=nil)
					if(incurrentcountlist[key]!=nil)
						data_product.each(){|pid,pname|
							inpaylist[key]+=(customlist[cid][current_date]['incurrentcountlist'][pid]*current_price[mostneardate][pid].to_f()-customlist[cid][current_date]['inwinningcountlist'][pid]*winning_price[mostneardate][pid].to_f())
						}
					else
						inpaylist[key]=0
					end
				else
					inpaylist[key]=0
				end	
			}

			
			#因為同一個GROUP的會被同時加進來，除以目前有的產品種類數就是正確的退水跟漲價金額了
			inaddmoney.each(){|key,value|
				customlist[cid][current_date]['inaddmoney']/=5
				customlist[cid][current_date]['inbonusmoney']/=5
			}
		

			row += inpaylist.collect(){|key,value| value}+[sumwitoutother,inaddmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x},inbonusmoney.collect(){|key,value| value.to_i()}.inject(){|sum,x| sum+x}]
		end
		sheet.write('A'+(8+rowindex).to_s(), row)
		rowindex+=1
	}


	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close


	#End
	p "#{free_current_date}_全產品4K當月月總計.xls 已輸出"

end







#客戶每日交易價格表
def CustomDailyPriceDetail(connection,current_date)
	#取得交易日期的年月日
	year=current_date.sub(/\/.*/,'')
	month=current_date.sub(year+'/','').sub(/\/.*/,'')
	day=current_date.sub(/.*\//,'')
	free_current_date=current_date.gsub(/\//,'')


	#取符合條件的產品
	recordset_product = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from product where CLng(PID)<100 order by CLng(PID);"
	recordset_product.Open(sql, connection)

	#取符合條件的使用者資料
	recordset_custom = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from custom;"
	recordset_custom.Open(sql, connection)

	#取符合條件的交易
	recordset_order = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from [order] where CLng(PID)<100 and CurrentDate='#{current_date}' order by group,CLng(PID);"
	recordset_order.Open(sql, connection)

	#取符合條件且最接近想搜尋的日期的一筆的價格
	recordset_price = WIN32OLE.new('ADODB.Recordset')
	sql = "select * from price where CLng(PID)<100 and CurrentDate='#{current_date}' order by CurrentDate desc,CLng(PID);"
	recordset_price.Open(sql, connection)






	#預存出所有會需要列出的資料
	#product
	data_product = recordset_product.GetRows.transpose
	pname='價格表'


	#custom
	data_custom = recordset_custom.GetRows.transpose


	#price
	data_price = recordset_price.GetRows.transpose
	current_price=Hash.new()
	winning_price=Hash.new()
	data_price.each(){|swiftcode,cid,pid,currentdate,currentprice,winningprice,upset|
		current_price[cid]=Hash.new() if current_price[cid]==nil
		current_price[cid][pid]=currentprice

		winning_price[cid]=Hash.new() if winning_price[cid]==nil
		winning_price[cid][pid]=winningprice
	}
	#倒序依客戶別排序
	current_price.sort_by{|key,v| key}
	winning_price.sort_by{|key,v| key}


	#order
	data_order = recordset_order.GetRows.transpose


	
	#Write the file
	FileUtils.mkdir_p("report/#{year}/#{month}/")
	# Create a new Excel Workbook
	book = WriteExcel.new("report/#{year}/#{month}/#{free_current_date}_客戶每日價格表.xls")
	# Add worksheet(s)
	sheet  = book.add_worksheet



	#Write a number and formula using A1 notation, and add row
	#顯示日期
	row = ['日期']
	format = book.add_format
	format.set_format_properties(:bg_color => 'gray',:pattern  => 1)
	sheet.write('A1',row, format)
	row = [current_date]
	sheet.merge_range('B1:P1',row, format)

	#列舉產品名
	pnamelist=Array.new
	data_product.each(){|pid,pname| 
		newpname=pname.sub(/.*_/,'') 
		pnamelist+=[newpname]
	}
	row = ['售價']+pnamelist
	sheet.write('A2', row, format)


	#列舉類別名
	row = ['成']
	sheet.write('A3', row, format)

	row = ['539']
	sheet.merge_range('B3:F3',row, format)

	row = ['港號']
	sheet.merge_range('G3:K3',row, format)

	row = ['大樂透']
	sheet.merge_range('L3:P3',row, format)

	
	#成本價
	rowindex=0
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=4+rowindex

		row=[cname]
		data_product.each(){|pid,pname|
			if(current_price[cid]==nil)
				row+=['0']
			else
				if(current_price[cid][pid]==nil)
					row+=['0']
				else
					row+=[current_price[cid][pid]]
				end
			end
		}

		sheet.write("A#{symbol}", row)
		rowindex+=1
	}


	#列舉類別名
	symbol=4+rowindex
	row = ['中']
	sheet.write("A#{symbol}", row, format)

	row = ['539']
	sheet.merge_range("B#{symbol}:F#{symbol}",row, format)

	row = ['港號']
	sheet.merge_range("G#{symbol}:K#{symbol}",row, format)

	symbol=4+rowindex
	row = ['大樂透']
	sheet.merge_range("L#{symbol}:P#{symbol}",row, format)
	rowindex+=1

	
	#中獎價
	data_custom.each(){|cid,cname,ctype,address,opendate,bankid,proportion,bonustarget,phone1,phone2,phone3,phone4,phone5,phone6,note|
		symbol=4+rowindex

		row=[cname]
		data_product.each(){|pid,pname|
			if(winning_price[cid]==nil)
				row+=['0']
			else
				if(winning_price[cid][pid]==nil)
					row+=['0']
				else
					row+=[winning_price[cid][pid]]
				end
			end
		}

		sheet.write("A#{symbol}", row)
		rowindex+=1
	}




	#close recordset
	recordset_product.close
	recordset_custom.close
	recordset_order.close
	recordset_price.close


	# write to file
	book.close

	#End
	p "#{free_current_date}_客戶每日價格表.xls 已輸出"

end





@connection = WIN32OLE.new('ADODB.Connection')
@connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=main.mdb')




case ARGV[0]
when 'CustomDailyTransactionDetail'
	#CustomDailyTransactionDetail(@connection,'100','1','2016/01/11')
	CustomDailyTransactionDetail(@connection,ARGV[1],ARGV[2],ARGV[3])
when 'DailyTransactionCounting'
	#DailyTransactionCounting(@connection,'100','2016/01/11')
	DailyTransactionCounting(@connection,ARGV[1],ARGV[2])
when 'AllDailyTransactionCounting'
	#AllDailyTransactionCounting(@connection,'2016/01/11')
	AllDailyTransactionCounting(@connection,ARGV[1])
when 'AllWeekTransactionCounting'
	#AllWeekTransactionCounting(@connection,'2016/01/11')
	AllWeekTransactionCounting(@connection,ARGV[1])
when 'AllMonthTransactionCounting'
	#AllMonthTransactionCounting(@connection,'2016/01/11')
	AllMonthTransactionCounting(@connection,ARGV[1])
when 'MonthTransactionCounting'
	#MonthTransactionCounting(@connection,'100','2016/01/11')
	MonthTransactionCounting(@connection,ARGV[1],ARGV[2])
when 'AllDaily4KTransactionCounting'
	#AllDaily4KTransactionCounting(@connection,'2016/01/11')
	AllDaily4KTransactionCounting(@connection,ARGV[1])
when 'AllMonth4KTransactionCounting'
	#AllMonth4KTransactionCounting(@connection,'2016/01/11')
	AllMonth4KTransactionCounting(@connection,ARGV[1])
when 'CustomDailyPriceDetail'
	#CustomDailyPriceDetail(@connection,'2016/01/11')
	CustomDailyPriceDetail(@connection,ARGV[1])
else
	p 'program error'
end