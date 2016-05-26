fout=File.open("./test.txt","w+")
	ARGV.each(){|a|		
		fout.write(a)
		fout.write("\n")
	}
fout.close()