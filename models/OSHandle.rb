# encoding: UTF-8


class OSHandle
  def initialize()
  end

  def finalize(object_id)  
  end

  def windows?
    (/cygwin|mswin|mingw|bccwin|wince|emx/ =~ RUBY_PLATFORM) != nil
  end

  def mac?
   (/darwin/ =~ RUBY_PLATFORM) != nil
  end

  def unix?
    !self.windows?
  end

  def linux?
    self.unix? and not self.mac?
  end
end