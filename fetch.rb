#!/usr/bin/env ruby
# encoding: utf-8

require 'rubygems'
require 'bundler/setup'

require 'pry'

require 'net/imap'
require 'mail'
require 'base64'
#require 'pdfkit'
require 'toml-rb'
require 'optparse'
require 'chrome_remote'
require 'base64'
require 'tempfile'

CONFIG_FILE_NAME = 'receipt-extractor.toml'
CONFIG_FILE_LOCATIONS = [
  "./#{CONFIG_FILE_NAME}",
  "~/.#{CONFIG_FILE_NAME}",
  "~/.receipt-extractor/#{CONFIG_FILE_NAME}",
  "~/.receipt-extractor/config.toml",
]

options = {
  mode: :mobility_package,
  config_file: CONFIG_FILE_LOCATIONS,
}
optparse = OptionParser.new do |opts|
  opts.banner = "Usage: $0 [options] [imap-filter]"

  opts.on("-c", "--config CONFIG", "Load configuration from CONFIG") do |v|
    options[:config_file] = [v]
  end

  opts.on('-e', '--expenses', "Switches to expenses mode (instead of mobility package mode). Currently only affects which FREE NOW payment methods are handled. Define the methods for each mode in the code in `PAYMENT_METHOD_MAP`.") do |v|
    options[:mode] = :expenses
  end

  opts.on('-h', '--handlers=HANDLERS', "Only run specific handlers. Pass a comma-separated list of handler IDs.") do |v|
    options[:only_handlers] = v.split(/,/).map(&:to_sym)
  end

  opts.on('-s', '--servers=SERVERS', "Only run on specific servers. Pass a comma-separated list of server keys (from config).") do |v|
    options[:only_servers] = v.split(/,/).map(&:to_sym)
  end

  opts.on('-f', '--force-overwrite', "Always overwrite files that exist (default off)") do |v|
    options[:force_overwrite] = true
    @force_overwrite = true
  end

  opts.on("--help", "Outputs this help") do
    STDERR.puts <<-_EOF_

Downloads receipts/invoices from IMAP servers and saves them to PDF files.

Syntax:

    $0 [--config CONFIG] [--expenses] [filter]

Arguments:

  -c CONFIG
  --config CONFIG
        Load configuration from CONFIG. Default search locations are:
#{CONFIG_FILE_LOCATIONS.map{ |x| "          #{x}" }.join("\n")}

  -e
  --expenses
        Switches to expenses mode (instead of mobility package mode). Currently
        only effects which FREE NOW payment methods are handled. Define the
        methods for each mode in the code in `PAYMENT_METHOD_MAP`.
  
  [filter]
        An IMAP filter to limit the timeframe to look at, or anything else.
        Is passed through to the IMAP server 1:1.
        Useful filters are:
          'SINCE 1-May-2018'
          'AFTER 1-Jan-2018'
          'FROM "name or email"'
          'SUBJECT "subject line"'
    _EOF_
    exit 99
  end
end
optparse.parse!

unless config_file = options[:config_file].map{ |f| File.expand_path f}.find{ |f| File.exists? f }
  STDERR.puts <<-_EOF_
Error: No configuration file found.
Run the command again with `--help' to see config file locations.
See `config.toml.sample` for an example.
_EOF_
  exit 1
end

begin
  @config = TomlRB.load_file(config_file, symbolize_keys: true)
  unless @config.has_key? :server then
    STDERR.puts <<-_EOF_
Error: Configuration file does not contain any [server.*] blocks.
See `config.toml.sample` for an example.
_EOF_
    exit 2
  end
  if @config[:server].any?{ |k,v| [:host,:port,:ssl,:username,:password].any?{ |x| !v.has_key? x } } then
    servers_with_errors = @config[:server].select{ |k,v| [:host,:port,:ssl,:username,:password].any?{ |x| !v.has_key? x} }.keys
    STDERR.puts <<-_EOF_
Error: The following server block(s) are incomplete: #{servers_with_errors.join(', ')}
See `config.toml.sample` for an example.
_EOF_
    exit 3
  end
  STDERR.puts "Loaded #{@config[:server].count} servers from configuration file: #{@config[:server].keys.join(', ')}"
rescue Exception => e
  STDERR.puts "#{e.class}: #{e.message}"
end

HANDLERS = {
  bvg: {
    label: "BVG",
    imap_filters: ['FROM "onlineshop@bvg.de"'],
    handler: :pdf_from_text
  },
  hvv: {
    label: "HVV",
    imap_filters: ['FROM "onlineshop@hochbahn.de"'],
    handler: :pdf_from_text
  },
  car2go: {
    label: "car2go",
    imap_filters: ['FROM "noreply@payment.car2go.com"'],
    handler: :special_car2go
  },
  mytaxi: {
    label: "FREE NOW / mytaxi",
    imap_filters: ['FROM "mytaxi Payment"', 'FROM "kundenservice@free-now.com"'],
    handler: :save_attachment,
    if: lambda { |message|
      payment_method = message.text_part.body.decoded.match(/Bezahlart: (.*)/)[1] rescue 'NULL'
      payment_method.gsub!(/ ,/, ',')
      payment_method.gsub!(/^\*(.*)\*$/, '\1')
      payment_method.gsub!(/ Trinkgeld.*/, '')
      return @config[:free_now][:payment_methods][@mode].include? payment_method
    }
  },
  coup: {
    label: "COUP",
    imap_filters: ['FROM "payment@joincoup.com"'],
    handler: :pdf_from_html
  },
  drivenow: {
    label: "Drive Now",
    imap_filters: ['SUBJECT "DriveNow eBilling"'],
    handler: :save_attachment
  },
  emmy: {
    label: "Emmy",
    imap_filters: ['SUBJECT "emmy Rechnung"'],
    handler: :save_attachment
  },
  uber: {
    label: "UBER",
    imap_filters: ['FROM "Uber Receipts" SUBJECT "trip receipt"'],
    handler: :pdf_from_html,
    if: lambda { |message|
      (message.header[:subject].decoded =~ /^\[Personal\]/ && @mode == :mobility_package) or
      (message.header[:subject].decoded =~ /^\[Business\]/ && @mode == :expenses) or
      (message.header[:subject].decoded =~ /Thanks for tipping\!.*trip receipt/ && @mode == :expenses)
    }
  },
  # uber_html: {
  #   label: "UBER save html",
  #   imap_filters: ['FROM "Uber Receipts" SUBJECT "trip receipt"'],
  #   handler: :save_html_part,
  #   if: lambda { |message|
  #     (message.header[:subject].decoded =~ /^\[Personal\]/ && @mode == :mobility_package) or
  #     (message.header[:subject].decoded =~ /^\[Business\]/ && @mode == :expenses) or
  #     (message.header[:subject].decoded =~ /Thanks for tipping\!.*trip receipt/ && @mode == :expenses)
  #   }
  # },
  deutsche_bahn: {
    label: "Deutsche Bahn",
    imap_filters: ['FROM "buchungsbestaetigung@bahn.de"', 'FROM "noreply.bahncard-rechnung@bahn.de"'],
    handler: :save_attachment
  },
  miles: {
    label: "MILES Sharing",
    imap_filters: ['SUBJECT "Deine MILES Rechnung"', 'SUBJECT "Deine drive by Rechnung"'],
    handler: :save_attachment
  },
  callabike: {
    label: "Call a Bike",
    imap_filters: ['SUBJECT "Call a Bike-Rechnung"'],
    handler: :save_attachment
  },
  hotel_invoice_marriott: {
    label: "Marriott Invoice",
    imap_filters: ['SUBJECT "Invoice of your stay"'],
    handler: :save_attachment
  },
  hotel_invoice_hilton: {
    label: "Hilton Invoice",
    imap_filters: ['FROM "receipt@hilton.com"'],
    handler: :save_attachment
  },
}

@mode = options[:mode]
@handlers = options[:only_handlers] || HANDLERS.keys

# this is the HTML inserted before an email to render it to PDF
TEXT_EMAIL_BEFORE = <<_EOF_
<html>
<head>
<style type="text/css">
html, body { font-size: 14px; font-family: Helvetica, Arial, sans-serif; }
hr { color: #dddddd; width: 100%; }
.mail-header th, td { font-size: 14px; }
.mail-header th { text-align: right; color: #888888; }
.mail-body { margin: 20px 20px; }
</style>
</head>
<body>
_EOF_

# this is the HTML inserted after an email to render it to PDF
TEXT_EMAIL_AFTER = <<_EOF_
</body></html>
_EOF_

# PDFKIT_OPTIONS = {
#   page_size: 'A4',
#   dpi: 300,
#   # disable_smart_shrinking: true,
#   # viewport_size: '500x1200',
#   margin_top: '15mm',
#   margin_left: '15mm',
#   margin_right: '15mm',
#   margin_bottom: '15mm',
# }

# PDFKit.configure do |config|
#   config.verbose = true
#   config.viewport_size = '720px'
# end

def get_messages(host, port, ssl, username, password)
  imap = Net::IMAP.new(host, port, ssl)
  imap.login(username, password)
  #puts imap.list('', '*')
  imap.select('INBOX')
  xf = ARGV.first+' ' rescue ''
  HANDLERS.slice(*@handlers).each do |id, data|
    STDERR.puts "- #{data[:label]} (#{id.to_s})"
    handler = data[:handler]
    handler_condition = data.fetch(:if, nil)
    data[:imap_filters].each do |filter|
      STDERR.puts "  - Filter: #{filter}" #" (==> #{'SINCE 1-Jan-2019 '+xf+filter})"
      uids = imap.uid_search('SINCE 1-Jan-2019 '+xf+filter)
      uids.each do |uid|
        message = Mail.new(imap.uid_fetch(uid, 'RFC822')[0].attr['RFC822'])
        puts "    - Message from #{message.header[:from].decoded}: \"#{message.header[:subject].decoded}\""
        condition = (handler_condition && handler_condition.call(message)) || !handler_condition
        if condition
          result = send "handler_"+handler.to_s, message
          puts "      #{handler.to_s.gsub(/_/, ' ')}"
          puts "        #{result.split(/\n/).join("\n        ")}" if result
        else
          puts "      (not matching condition, skipped)"
        end
      end
    end
  end
  imap.logout
  imap.disconnect
end

def handler_skip(message); end
def handler_unhandled(message); binding.pry; end

def handler_pdf_from_text(message)
  filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
  return "(already exists, skipping) #{filename}" if File.exist?(filename) and not @force_overwrite
  chrome_render_pdf(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + helper_text_to_html(message.decoded) + TEXT_EMAIL_AFTER, filename)
  return "(saved) "+filename
end

def handler_pdf_from_html(message)
  filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
  return "(already exists, skipping) #{filename}" if File.exist?(filename) and not @force_overwrite
  body = message.html_part.decoded
  if body.match(/cid:/)
    # embedded files
    message.parts.select(&:content_id).each do |part|
      cid = part.content_id[1..-2]
      body.gsub!(/cid:#{cid}/, "data:#{part.content_type};base64,#{Base64::strict_encode64(part.decoded)}")
    end
  end
  chrome_render_pdf(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + body + TEXT_EMAIL_AFTER, filename)
  return "(saved) "+filename
end

# def handler_pdf_from_text(message)
#   kit = PDFKit.new(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + helper_text_to_html(message.decoded) + TEXT_EMAIL_AFTER, PDFKIT_OPTIONS)
#   filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
#   return "(already exists, skipping) #{filename}" if File.exist?(filename) and not @force_overwrite
#   # while File.exist?(filename)
#   #   counter = filename.match(/_(\d+)\.pdf/)[1] rescue 0
#   #   counter += 1
#   #   filename = filename.sub(/(?:_\d+)?\.pdf/, "_#{counter}.pdf")
#   # end
#   kit.to_file(filename)
#   return "(saved) "+filename
# end

# def handler_pdf_from_html(message)
#   body = message.html_part.decoded
#   if body.match(/cid:/)
#     # embedded files
#     message.parts.select(&:content_id).each do |part|
#       cid = part.content_id[1..-2]
#       body.gsub!(/cid:#{cid}/, "data:#{part.content_type};base64,#{Base64::strict_encode64(part.decoded)}")
#     end
#   end
#   kit = PDFKit.new(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + body + TEXT_EMAIL_AFTER, PDFKIT_OPTIONS)
#   filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
#   return "(already exists, skipping) #{filename}" if File.exist?(filename) and not @force_overwrite
#   # while File.exist?(filename)
#   #   counter = filename.match(/_(\d+)\.pdf/)[1] rescue 0
#   #   counter += 1
#   #   filename = filename.sub(/(?:_\d+)?\.pdf/, "_#{counter}.pdf")
#   # end
#   kit.to_file(filename)
#   return "(saved) "+filename
# end

def handler_special_car2go(message)
  pdf_attachments = message.attachments.select{ |x| x.content_type.start_with? 'application/pdf' }
  return "(skipping) no attachments in message: #{message.header[:from].decoded}: #{message.header[:subject].decoded}" if pdf_attachments.count == 0
  return "(skipping) Lastschriftvorankündigung" if message.header[:subject].decoded =~ /Lastschriftvorankündigung/
  if message.header[:subject].decoded !~ /Deine neue Rechnung/
    puts message.header[:subject].decoded
    binding.pry
  end
  r = []
  pdf_attachments.each do |attachment|
    filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}__#{attachment.filename.sub(/\.[^.]+$/, '')}.pdf"
    if File.exist?(filename) and not @force_overwrite
      r << "(already exists) #{filename}"
    else
      begin
        File.open(filename, "w+b", 0644) {|f| f.write attachment.decoded}
        r << "(saved) "+filename
      rescue => e
        r << "(ERROR) Unable to save data for #{filename} because #{e.message}"
        STDERR.puts "(ERROR) Unable to save data for #{filename} because #{e.message}"
      end
    end
  end
  r.join("\n")
end

def handler_save_attachment(message)
  return "(skipping) no attachments in message: #{message.header[:from].decoded}: #{message.header[:subject].decoded}" if message.attachments.count == 0
  r = []
  message.attachments.each do |attachment|
    filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}__#{attachment.filename.sub(/\.[^.]+$/, '')}.pdf"
    if File.exist?(filename) and not @force_overwrite
      r << "(already exists) #{filename}"
    else
      begin
        File.open(filename, "w+b", 0644) {|f| f.write attachment.decoded}
        r << "(saved) "+filename
      rescue => e
        r << "(ERROR) Unable to save data for #{filename} because #{e.message}"
        STDERR.puts "(ERROR) Unable to save data for #{filename} because #{e.message}"
      end
    end
  end
  r.join("\n")
end

def handler_save_html_part(message)
  r = []
  filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.html"
  if File.exist?(filename) and not @force_overwrite
    r << "(already exists) #{filename}"
  else
    begin
      body = message.html_part.decoded
      if body.match(/cid:/)
        # embedded files
        message.parts.select(&:content_id).each do |part|
          cid = part.content_id[1..-2]
          body.gsub!(/cid:#{cid}/, "data:#{part.content_type};base64,#{Base64::strict_encode64(part.decoded)}")
        end
      end
      File.open(filename, "w+b", 0644) do |f|
        f.write TEXT_EMAIL_BEFORE
        f.write helper_mail_header_html(message)
        f.write "\n\n<!-- - - - - - - - - - - - - - - - - - - - - - - -->\n\n"
        f.write body
        f.write TEXT_EMAIL_AFTER
      end
      r << "(saved) "+filename
    rescue => e
      r << "(ERROR) Unable to save data for #{filename} because #{e.message}"
      STDERR.puts "(ERROR) Unable to save data for #{filename} because #{e.message}"
    end
  end
  r.join("\n")
end

def helper_mail_header_html(message)
  _ = <<_EOF_
<div class="mail-header"><table border="0" width="100%">
<tr><th>From:</b></th> <td width="100%">#{message.header[:from].decoded}</td></tr>
<tr><th>Subject:</th> <td>#{message.header[:subject].decoded}<br></td></tr>
<th>Date:</th> <td>#{message.date.strftime('%Y-%m-%d %H:%M')}<br></td></tr>
<th>To:</th> <td>#{message.header[:to].decoded}</td></tr>
</table></div>
<hr noshade="noshade">
_EOF_
end

def helper_text_to_html(text)
  '<div class="mail-body"><p>'+text.split(/\n\n/).join('</p><p>').split(/\n/).join('<br>')+'</p></div>'
end

def init_chrome
  return if @chrome
  port = rand(15222..25222)
  chrome_pid = spawn("/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --headless --remote-debugging-port=#{port} --no-first-run --hide-scrollbars --disable-gpu --disable-sync --disable-translate >/dev/null 2>/dev/null &")
  @chrome = nil
  at_exit {
    begin
      @chrome.send_cmd "Browser.close"
      Process.wait chrome_pid
    rescue Exception => e
      # pass
    end
  }
  while not @chrome = ChromeRemote.client(port: port) rescue nil do
    sleep 0.25
  end
  @chrome.send_cmd "Page.enable"
end

def chrome_render_pdf(content, outfile)
  init_chrome
  Tempfile.create([outfile, '.html']) do |f|
    $stderr.puts f.path
    f.write(content)
    f.close
    @chrome.send_cmd "Page.navigate", url: "file://#{f.path}", referer: 'file://'
    @chrome.wait_for "Page.loadEventFired"
    @chrome.send_cmd "Emulation.setEmulatedMedia", media: 'screen'
    response = @chrome.send_cmd "Page.printToPDF", pageSize: 'A4', marginTop: 0, printBackground: true, marginBottom: 0, marginLeft: 0, marginRight: 0, displayHeaderFooter: true, scale: 0.7, transferMode: 'ReturnAsBase64'
    File.write outfile, Base64.decode64(response['data'])
  end
end

options[:only_servers] ||= @config[:server].keys
@config[:server].slice(*options[:only_servers]).each do |name, server|
  STDERR.puts "Fetching from server '#{name}'..."
  get_messages server[:host], server[:port], server[:ssl], server[:username], server[:password]
end
