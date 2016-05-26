require 'FileUtils'

ARGV.each(){|a|
	p a
	if(a=='test')
		FileUtils.mkdir_p('./test')
	end
}