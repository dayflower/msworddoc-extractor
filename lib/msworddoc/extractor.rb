require 'rubygems'
require 'ole/storage'

module MSWordDoc
  module Extractor
    VERSION = '0.1.0'

    def self.load(file)
      doc = Essence.new()

      ole = Ole::Storage.open(file)

      doc.load_storage(ole)

      if block_given?
        begin
          yield doc
        ensure
          doc.close()
        end

        return
      end

      return doc
    end
  end

  class Essence
    PPS_NAME_WORDDOC = 'WordDocument'
    PPS_NAME_TABLE_TMPL = '%dTable'

    MAGIC_MSWORD = 0xa5ec
    NFIB_MSWORD6 = 101

    OFFSET_FIB_IDENT = 0x0000
    OFFSET_FIB_FIB   = 0x0002

    OFFSET_FIB_FLAGS  = 0x000a
    OFFSET_FIB_FCCLX  = 0x01a2
    OFFSET_FIB_LCBCLX = 0x01a6

    OFFSET_FIB_FCMIN = 0x0018
    OFFSET_FIB_FCMAC = 0x001c
    OFFSET_FIB_CBMAC = 0x0040

    MASK_FIBFLAG_COMPLEX     = 0x0004
    MASK_FIBFLAG_ENCRYPTED   = 0x0100
    MASK_FIBFLAG_WHICHTBLSTM = 0x0200

    LENGTH_CP  = 4
    LENGTH_PCD = 8

    OFFSET_FIB_CCP_MAP = {
      :Text    => 0x004c,
      :Ftn     => 0x0050,
      :Hdd     => 0x0054,
      :Mcr     => 0x0058,
      :Atn     => 0x005c,
      :Edn     => 0x0060,
      :Txbx    => 0x0064,
      :HdrTxbx => 0x0068,
    }

    def initialize
      @flag = {}
      @ccp  = {}
      @ole = nil
    end

    def close
      @ole.close()
      @ole = nil
    end

    def load_storage(ole)
      @ole = ole

      @ole.file.open(PPS_NAME_WORDDOC) do |f|
        parse_fib(f)
      end

      name_of_table = PPS_NAME_TABLE_TMPL % (@flag[:fWhichTblStm] ? 1 : 0)
      @ole.file.open name_of_table do |f|
        parse_piece_table(f)
      end
    end

    def whole_contents(*args)
      return retrieve_and_filter(0, -1, *args)
    end

    def document(*args)
      return retrieve_and_filter(0, @ccp[:Text], *args)
    end

    def footnote(*args)
      return retrieve_and_filter(@ccp[:Text], @ccp[:Ftn], *args)
    end

    def header(*args)
      skips = [ :Text, :Ftn ]
      return retrieve_token_and_filter(skips, :Hdd, *args)
    end

    def macro(*args)
      skips = [ :Text, :Ftn, :Hdd ]
      return retrieve_token_and_filter(skips, :Mcr, *args)
    end

    def annotation(*args)
      skips = [ :Text, :Ftn, :Hdd, :Mcr ]
      return retrieve_token_and_filter(skips, :Atn, *args)
    end

    def endnote(*args)
      skips = [ :Text, :Ftn, :Hdd, :Mcr, :Atn ]
      return retrieve_token_and_filter(skips, :Edn, *args)
    end

    def textbox(*args)
      skips = [ :Text, :Ftn, :Hdd, :Mcr, :Atn, :Edn ]
      return retrieve_token_and_filter(skips, :Txbx, *args)
    end

    def header_textbox(*args)
      skips = [ :Text, :Ftn, :Hdd, :Mcr, :Atn, :Edn, :Txbx ]
      return retrieve_token_and_filter(skips, :HdrTxbx, *args)
    end

    private

    def parse_fib(f)
      if get_ushort(f, OFFSET_FIB_IDENT) != MAGIC_MSWORD then
        raise 'Not a Word document'
      end

      nFib = get_ushort(f, OFFSET_FIB_FIB)
      if nFib < NFIB_MSWORD6
        raise 'Unsupported version'
      end

      flags = get_ushort(f, OFFSET_FIB_FLAGS)

      @flag[:fComplex] = (flags & MASK_FIBFLAG_COMPLEX != 0)

      @flag[:fEncrypted] = (flags & MASK_FIBFLAG_ENCRYPTED != 0)
      if @flag[:fEncypted]
        raise 'Encrypted MSWord document file is not supported'
      end

      @flag[:fWhichTblStm] = (flags & MASK_FIBFLAG_WHICHTBLSTM != 0)

      @fcMin = get_ulong(f, OFFSET_FIB_FCMIN)
      @fcMac = get_ulong(f, OFFSET_FIB_FCMAC)
      @cbMac = get_ulong(f, OFFSET_FIB_CBMAC)

      @fcClx  = get_ulong(f, OFFSET_FIB_FCCLX)
      @lcbClx = get_ulong(f, OFFSET_FIB_LCBCLX)

      parse_fib_ccps(f)
    end

    def parse_fib_ccps(f)
      OFFSET_FIB_CCP_MAP.each do |key, offset|
        @ccp[key] = get_ulong(f, offset)
      end
    end

    def parse_piece_table(f)
      if @lcbClx <= 0
        # create pseudo piece table
        ccpAll = 0
        OFFSET_FIB_CCP_MAP.each do |key, offset|
          ccpAll += @ccp[key]
        end

        @pcds = [
          {
            :fc  => @fcMin,
            :cp  => 0,
            :ccp => ccpAll,
          }
        ]

        return
      end

      f.pos = @fcClx
      clx = f.read(@lcbClx)

      while clx.length > 0
        clxt = clx.slice!(0, 1).unpack('C')[0]
        break if clxt == 2        # plcfpcd

        if clxt == 1              # grpprl => SKIP
          skip = clx.slice!(0, 2).unpack('v')[0]

          clx.slice!(0, skip)
        else
          raise 'Unknown CLX block'
        end
      end
      raise 'PCDs not found'  unless clx.length > 0

      length = clx.slice!(0, 4).unpack('V')[0]

      n = (length - LENGTH_CP) / (LENGTH_CP + LENGTH_PCD)

      cps = []
      (n+1).times do
        cps << clx.slice!(0, LENGTH_CP).unpack('V')[0]
      end

      @pcds = []
      1.upto(n) do |i|
        pcd_data = clx.slice!(0, LENGTH_PCD)

        fc = pcd_data.slice(2, 4).unpack('V')[0]

        @pcds << {
          :fc  => fc,
          :cp  => cps[i - 1],
          :ccp => cps[i] - cps[i - 1]
        }
      end
    end

    def retrieve_substring(f, offset, length = -1)
      i = 0
      while i < @pcds.length
        if @pcds[i][:cp] > offset then break end

        i += 1
      end
      i -= 1
      raise 'could not find suitable heading piece' unless i >= 0

      output = ""
      while length > 0 || length < 0
        pcd = @pcds[i]

        len = length
        if pcd[:ccp] < len || len < 0
          len = pcd[:ccp]
        end

        if pcd[:fc] & 0x40000000 != 0
          # cp1252
          fc = (pcd[:fc] ^ 0x40000000) >> 1
          fc += offset
          offset = 0

          f.pos = fc
          output << convert_from_cp1252(f.read(len))
        else
          # UTF-16LE
          fc = pcd[:fc]
          fc += offset * 2
          offset = 0

          f.pos = fc
          output << convert_from_utf16le(f.read(len * 2))
        end

        if length >= 0
          length -= len
        end

        i += 1
        break if i >= @pcds.length
      end

      return output
    end

    def get_ushort(f, pos)
      f.pos = pos
      return f.read(2).unpack('v')[0]
    end

    def get_ulong(f, pos)
      f.pos = pos
      return f.read(4).unpack('V')[0]
    end

    def retrieve_token_and_filter(skip_tokens, target, *args)
      skip = skip_tokens.inject(0) {|sum, key| sum + @ccp[key] }
      return retrieve_and_filter(skip, @ccp[target], *args)
    end

    def retrieve_and_filter(offset, length, *args)
      opts = Hash[*args]

      string = ""
      @ole.file.open PPS_NAME_WORDDOC do |f|
        string = retrieve_substring(f, offset, length)
      end

      if ! opts[:raw]
        return format_into_plain(string)
      end

      return string
    end

    CHARMAP = {
      "\x0d" => "\n",         # ASIS: Line Feed
      "\x09" => "\t",         # ASIS: Tab

      "\x0d" => "\n",         # Paragraph ends; \n + U+2029?

      "\x0b" => "\n",         # Hard line breaks

      "\x2d" => "\x2d",       # ASIS: Breaking hyphens; U+2010?
      "\x1f" => "\u{00ad}",   # Non-required hyphens (into Soft hyphen)
      "\x1e" => "\u{2011}",   # Non-breaking hyphens

      "\xa0" => "\xa0",       # ASIS: Non-breaking-spaces

      "\x0c" => "\x0c",       # ASIS: Page breaks or Section marks

      "\x0e" => "\x0e",       # ASIS: Column breaks

      "\x13" => "",           # Field begin mark
      "\x15" => "",           # Field end mark
      "\x14" => "",           # Field separator

      "\x07" => "\t",         # Cell mark or Row mark
    }

    def format_into_plain(text)
      text = text.gsub(/([\x07]*)[\x07]{2}/, '\1\n') \
                 .gsub(/([\x00-\x1f])/) { CHARMAP[$1] || "" }

      return text
    end

    if defined?(Encoding)
      # for Ruby 1.9+

      def convert_from_cp1252(str)
        @enc_utf8   ||= Encoding.find('UTF-8')
        @enc_cp1252 ||= Encoding.find('Windows-1252')
        return str.encode(@enc_utf8, @enc_cp1252)
      end

      def convert_from_utf16le(str)
        @enc_utf8  ||= Encoding.find('UTF-8')
        @enc_utf16 ||= Encoding.find('UTF-16LE')
        return str.encode(@enc_utf8, @enc_utf16)
      end
    else
      # for Ruby 1.8
      require 'nkf'

      def convert_from_cp1252(str)
        return NKF.nkf(dest_encoding() + ' -W', str)
      end

      def convert_from_utf16le(str)
        return NKF.nkf(dest_encoding() + ' -W16L0', str)
      end

      def dest_encoding
        case $KCODE
        when /^E/i then '-e'
        when /^S/i then '-s'
        when /^U/i then '-w'
        else            '-w'
        end
      end
    end

  end
end
