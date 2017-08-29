require 'csv'
require 'pp'
require 'readline'
require 'docx'
require 'zip'
require 'cgi'


NEWLINE = '</w:t></w:r></w:p><w:p w:rsidR="00CC4A7B" w:rsidRDefault="00CC4A7B" w:rsidP="00CC4A7B"><w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">    </w:t></w:r><w:r><w:t>'
NEWBOLLET = '</w:t></w:r></w:p><w:p w:rsidR="00D923D3" w:rsidRDefault="00CE7F67" w:rsidP="00512F22"><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:bidi w:val="0"/></w:pPr><w:r><w:t>'
TABLEROW_HIGH = 	'<w:tr w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidTr="007F70C0"><w:trPr><w:jc w:val="center"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1418" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="0050148A" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/><w:rPr><w:iCs/></w:rPr></w:pPr><w:r><w:rPr><w:iCs/></w:rPr><w:t>MRDCVE</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="00FE7B6E" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r w:rsidRPr="00C862C2"><w:rPr><w:noProof/><w:lang w:bidi="ar-SA"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="1227D5C3" wp14:editId="1D406479"><wp:extent cx="194323" cy="136026"/><wp:effectExtent l="0" t="0" r="15240" b="16510"/><wp:docPr id="134" name="Rounded Rectangle 134"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr><a:spLocks/></wps:cNvSpPr><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="194323" cy="136026"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent6"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val="accent6"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent6"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent6"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></wps:style><wps:txbx><w:txbxContent><w:p w:rsidR="00FE7B6E" w:rsidRDefault="00FE7B6E" w:rsidP="00FE7B6E"/></w:txbxContent></wps:txbx><wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="1" fromWordArt="0" anchor="ctr" anchorCtr="0" forceAA="0" compatLnSpc="1"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:roundrect id="Rounded Rectangle 134" o:spid="_x0000_s1026" style="width:15.3pt;height:10.7pt;visibility:visible;mso-wrap-style:square;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;mso-left-percent:-10001;mso-top-percent:-10001;v-text-anchor:middle" arcsize="10923f" o:gfxdata="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&#xA;" fillcolor="#e36c0a [2409]" strokecolor="#e36c0a [2409]" strokeweight="2pt"><v:path arrowok="t"/><v:textbox><w:txbxContent><w:p w:rsidR="00FE7B6E" w:rsidRDefault="00FE7B6E" w:rsidP="00FE7B6E"/></w:txbxContent></v:textbox><w10:anchorlock/></v:roundrect></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2608" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="0050148A" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="C00000"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="mb5"/><w:b w:val="0"/></w:rPr><w:t>MRDCVSS</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="936" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="00577B5A" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="cs"/><w:rtl/></w:rPr><w:t>زیاد</w:t></w:r></w:p></w:tc></w:tr>'
TABLEROW_CRITICAL = '<w:tr w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidTr="007F70C0"><w:trPr><w:jc w:val="center"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1418" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="0050148A" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/><w:rPr><w:iCs/></w:rPr></w:pPr><w:r><w:rPr><w:iCs/></w:rPr><w:t>MRDCVE</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="00FE7B6E" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r w:rsidRPr="00C862C2"><w:rPr><w:noProof/><w:lang w:bidi="ar-SA"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="4C550687" wp14:editId="74E84D25"><wp:extent cx="194323" cy="136026"/><wp:effectExtent l="0" t="0" r="15240" b="16510"/><wp:docPr id="134" name="Rounded Rectangle 134"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr><a:spLocks/></wps:cNvSpPr><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="194323" cy="136026"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent6"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent6"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></wps:style><wps:txbx><w:txbxContent><w:p w:rsidR="00FE7B6E" w:rsidRDefault="00FE7B6E" w:rsidP="00FE7B6E"/></w:txbxContent></wps:txbx><wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="1" fromWordArt="0" anchor="ctr" anchorCtr="0" forceAA="0" compatLnSpc="1"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:roundrect id="Rounded Rectangle 134" o:spid="_x0000_s1026" style="width:15.3pt;height:10.7pt;visibility:visible;mso-wrap-style:square;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;mso-left-percent:-10001;mso-top-percent:-10001;v-text-anchor:middle" arcsize="10923f" o:gfxdata="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&#xA;cmV2LnhtbFBLBQYAAAAABAAEAPMAAAD7BQAAAAA=&#xA;" fillcolor="red" strokecolor="red" strokeweight="2pt"><v:path arrowok="t"/><v:textbox><w:txbxContent><w:p w:rsidR="00FE7B6E" w:rsidRDefault="00FE7B6E" w:rsidP="00FE7B6E"/></w:txbxContent></v:textbox><w10:anchorlock/></v:roundrect></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2608" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="0050148A" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="C00000"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="mb5"/><w:b w:val="0"/></w:rPr><w:t>MRDCVSS</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="936" w:type="dxa"/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00C862C2" w:rsidRDefault="00E77497" w:rsidP="007F70C0"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="cs"/><w:rtl/></w:rPr><w:t>بحرانی</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p></w:tc></w:tr>'
TABLEROW_HOSTS = '<w:tr w:rsidR="00FE7B6E" w:rsidRPr="00741B63" w:rsidTr="007A1552"><w:trPr><w:trHeight w:val="300"/><w:jc w:val="center"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1610" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:noWrap/><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00741B63" w:rsidRDefault="0050148A" w:rsidP="00561817"><w:pPr><w:bidi w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>MRDIPS</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1797" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:noWrap/><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00741B63" w:rsidRDefault="00561817" w:rsidP="00561817"><w:pPr><w:bidi w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>MRDPRT</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1296" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="FF0000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:noWrap/><w:vAlign w:val="center"/><w:hideMark/></w:tcPr><w:p w:rsidR="00FE7B6E" w:rsidRPr="00741B63" w:rsidRDefault="0050148A" w:rsidP="00561817"><w:pPr><w:bidi w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>MRDPORTS</w:t></w:r></w:p></w:tc></w:tr>'
BODY_START = '<w:body>'
BODY_END = '<w:sectPr w:rsidR="00D923D3">'


def x bind
        puts "Execuation mode: "
        begin
                begin
                  s = Readline.readline("exe> ",true).strip
                  eval ("puts ' = ' + (#{s}).to_s"),bind
                rescue => e
                  puts " > Error ocurred: #{e.backtrace[0]}: #{e.message}"
                end while true
        ensure
                nil
        end
end
 

@data = []
Dir.new('.').each {|file| 
	if file.end_with?("csv") 
		CSV.foreach(file,:headers => true, :header_converters => :symbol)do |row|
			@data << row.to_hash
		end
	end
}

def psearch(text)
   pp search(text)
   nil
end

def phosts(pid)
   x= hosts(pid.to_s)
   pp x.map{|row| "#{row[0]} #{row[1]} #{row[2]}            " }
   nil
end

def search(text)
  @data.select{|row| row[:risk]=="Critical" && row[:name].downcase.include?(text.downcase) }.map{|z| [z[:plugin_id],z[:name]]}.uniq
end

def hosts(pid)
  @data.select{|row| row[:plugin_id]==pid }.map{|z| [z[:host],z[:port],z[:protocol]]}.uniq
end

def prisks
   pp @data.map{|row| row[:risk] }.to_a.uniq
   nil
end

def byrisk risk
   @data.select{|row| row[:risk]==risk }.sort_by{|z| z[:cvss]}.map{|z| [z[:plugin_id],z[:name]]}.sort_by{|z| z[0]}.to_a.uniq
end

def pbyrisk risk
   pp byrisk risk
   nil
end

def pcve pid
	pp @data.select{|z| z[:plugin_id] == pid.to_s}.map{|a| "#{a[:cve]} #{a[:cvss]}    "}.sort_by{|a| a[:cve]}.uniq
	nil
end


def psynopsis pid
	puts @data.select{|z| z[:plugin_id] == pid.to_s}[0][:synopsis]
	nil
end

def pdescription pid
	puts @data.select{|z| z[:plugin_id] == pid.to_s}[0][:description]
	nil
end

def psolution pid
	puts @data.select{|z| z[:plugin_id] == pid.to_s}[0][:solution]
	nil
end

def finishit
	export2word "Critical","./temp - Critical.docx"
	export2word "High","./temp - High.docx"
end

def export2word lvl,file_path
    cntt = ''
	byrisk(lvl).each{|a| 
		cntt += toword a,file_path
	}	
	zip_file_name = "./#{lvl}.docx"
	FileUtils.cp file_path,zip_file_name
	Zip::File.open(zip_file_name) do |zipfile|
	  files = zipfile.select(&:file?)
	  files.each do |zip_entry|
		if(zip_entry.name == "word/document.xml")
			s = zipfile.read(zip_entry.name).force_encoding("utf-8")
			start__ = s.index(BODY_START) + BODY_START.size
			end__ = s.index(BODY_END)
			# cntt
			s2 = s[0...start__] +  cntt  + s[end__..-1]
			zipfile.get_output_stream(zip_entry.name){ |f| f.puts s2 }
		end
	  end
	  zipfile.commit
	end
	
end

def toword a,file_path
    result = ''
	#zip_file_name = "./#{a[0]}.docx"
	#FileUtils.cp file_path,zip_file_name
	row_cve = TABLEROW_HIGH
	if(file_path.include? "Critical")
		row_cve = TABLEROW_CRITICAL
	end
	title = a[1]
	cvess = @data.select{|z| z[:plugin_id] == a[0].to_s}.map{|a| [a[:cve],a[:cvss]]}.sort_by{|z| z[0]}.uniq
	hosts = @data.select{|row| row[:plugin_id]==a[0] }.map{|z| [z[:host],z[:port],z[:protocol]]}.uniq
	desc = @data.select{|z| z[:plugin_id] == a[0].to_s}[0][:synopsis] + "\n" + @data.select{|z| z[:plugin_id] == a[0].to_s}[0][:description]
	solution = @data.select{|z| z[:plugin_id] == a[0].to_s}[0][:solution]
	plugin_out = @data.select{|z| z[:plugin_id] == a[0].to_s}[0][:plugin_output]
	see_also = @data.select{|z| z[:plugin_id] == a[0].to_s}[0][:see_also]
	cve_rows = cvess.map{|a|  row_cve.gsub("MRDCVE",CGI.escapeHTML(a[0])).gsub("MRDCVSS",CGI.escapeHTML(a[1]))  }.join
	hosts_rows = hosts.map{|a|  TABLEROW_HOSTS.gsub("MRDIPS",CGI.escapeHTML(a[0])).gsub("MRDPRT",CGI.escapeHTML(a[2])).gsub("MRDPORTS",CGI.escapeHTML(a[1]))  }.join
	
	Zip::File.open(file_path) do |zipfile|
	  files = zipfile.select(&:file?)
	  files.each do |zip_entry|
		if(zip_entry.name == "word/document.xml")
			s = zipfile.read(zip_entry.name).force_encoding("utf-8")
			s.gsub!("MRDTITLE",CGI.escapeHTML(title) )
			s.gsub!(row_cve,cve_rows)
			s.gsub!(TABLEROW_HOSTS,hosts_rows)
			s.gsub!("MRDDESCRIPTION",CGI.escapeHTML(desc).gsub("\n",NEWLINE))
			s.gsub!("MRDSOLUTION",CGI.escapeHTML(solution).gsub("\n",NEWLINE))
			s.gsub!("MRDOUTPUT",CGI.escapeHTML(plugin_out).gsub("\n",NEWLINE).gsub("\r",""))
			s.gsub!("MRDSEEALSO",CGI.escapeHTML(see_also).gsub("\n",NEWBOLLET))
				
				
			start__ = s.index(BODY_START) + BODY_START.size
			end__ = s.index(BODY_END)
			
			result =s[start__...end__]
		end
	  end
	  zipfile.commit
	end
	return result
end


def doloop level
	byrisk(level).each{|a| 

		puts "------------------------------------------"
		puts a[1]
		puts "*******************"
		puts "CVE:"
		pcve a[0]
		puts "*******************"
		puts "Synopsis:"
		psynopsis a[0]
		puts "*******************"
		puts "Description:"
		pdescription a[0]
		puts "*******************"
		puts "Hosts:"
		phosts a[0]
		puts "*******************"
		puts "Solution:"
		psolution a[0]
		puts "*******************"
		puts "Continue..."
		gets
	}
end

x binding
# 
# 