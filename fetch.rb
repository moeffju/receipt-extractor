#!/usr/bin/env ruby
# encoding: utf-8

require 'rubygems'
require 'bundler/setup'

require 'pry'

require 'net/imap'
require 'mail'
require 'base64'
require 'pdfkit'

if ARGV.first == '--help' || ARGV.first == '-h'
  STDERR.puts <<_EOF_
Downloads receipts from IMAP servers and saves them to PDF files.

Syntax:

    $0 [--expenses] [filter]

Arguments:

  --expenses
        Switches to expenses mode (instead of mobility package mode). Currently
        only effects which mytaxi payment methods are handled. Define the
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
  exit 1
end

@mode = :mobility_package
@mode = :expenses if ARGV.shift == '--expenses'

# list your servers here. this should go into a config file.
SERVERS = [
  {host: 'imap.gmail.com', port: 993, ssl: true, username: 'user@gmail.com', password: 'password'},
  {host: 'mail.example.com', port: 143, ssl: false, username: 'user', password: 'password'},
]

# this is a list of IMAP filters that will be stepped through in order, combined with any filters given on the commandline
FILTERS = <<_FILTERS_.split("\n").map(&:strip)
  FROM "onlineshop@bvg.de"
  FROM "onlineshop@hochbahn.de"
  FROM "noreply@payment.car2go.com"
  FROM "payment@joincoup.com"
  SUBJECT "DriveNow eBilling"
  SUBJECT "emmy Rechnung"
  FROM "mytaxi Payment"
  SUBJECT "trip with Uber"
  FROM "buchungsbestaetigung@bahn.de"
  SUBJECT "Deine MILES Rechnung"
  SUBJECT "Deine drive by Rechnung"
  SUBJECT "Call a Bike-Rechnung"
  FROM "noreply.bahncard-rechnung@bahn.de"
_FILTERS_

# this is the HTML inserted before a text email to render it to PDF
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

# this is the HTML inserted after a text email to render it to PDF
TEXT_EMAIL_AFTER = <<_EOF_
</body></html>
_EOF_

PAYMENT_METHOD_MAP = {
  mytaxi: {
    mobility_package: ['Bar', 'Kreditkarte, 123456******9876', 'Kreditkarte, 654321******6789'],
    expenses: ['Kreditkarte, 123123******1234', 'Kreditkarte, 123123******2345', nil],
  },
}

def get_messages(host, port, ssl, username, password)
  imap = Net::IMAP.new(host, port, ssl)
  imap.login(username, password)
  #puts imap.list('', '*')
  imap.select('INBOX')
  xf = ARGV.first+' ' rescue ''
  FILTERS.each do |filter|
    uids = imap.uid_search('SINCE 1-Jan-2019 '+xf+filter)
    uids.each do |uid|
      message = Mail.new(imap.uid_fetch(uid, 'RFC822')[0].attr['RFC822'])
      handle_message(message)
    end
  end
  imap.logout
  imap.disconnect
end

# add cases here
def handle_message(message)
  handler = :unhandled
  from = message.from[0].to_s.downcase
  case from
  when 'onlineshop@bvg.de', 'onlineshop@hochbahn.de'
    # create pdf from text
    handler = :pdf_from_text
  when 'payment@joincoup.com'
    # create pdf from html
    handler = :pdf_from_html
  when 'noreply@payment.car2go.com'
    # check if Lastschriftvorankündigung
    # extract attachment
    # extract page 2
    handler = :special_car2go
  when 'service@drive-now.com', 'info@emmy-sharing.de', 'buchungsbestaetigung@bahn.de', 'noreply.bahncard-rechnung@bahn.de', 'hello@miles-mobility.com', 'info@callabike.de'
    # extract attachment
    handler = :save_attachment
  else
    if from =~ /uber\..*@uber.com$/
      # create pdf from html
      handler = :pdf_from_html
    elsif from =~ /@mytaxi.com$/
      mytaxi_payment_method = message.text_part.body.decoded.match(/Bezahlart: \*(.*)\*/)[1] rescue nil
      case mytaxi_payment_method
      when PAYMENT_METHOD_MAP[:mytaxi][@mode]
        handler = :save_attachment
      else
        handler = :skip
      end
    elsif from =~ /@drive-by\.de/
      handler = :save_attachment
    end
  end
  puts "[#{handler.to_s.gsub(/_/, ' ')}] #{message.header[:from].decoded}: #{message.header[:subject].decoded}"
  result = send "handler_"+handler.to_s, message
  print "    #{result.split(/\n/).join("\n    ")}\n" if result
end

def handler_skip(message); end
def handler_unhandled(message); binding.pry; end

def handler_pdf_from_text(message)
  kit = PDFKit.new(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + helper_text_to_html(message.decoded) + TEXT_EMAIL_AFTER, page_size: 'A4', margin_top: '15mm', margin_left: '15mm', margin_right: '15mm', margin_bottom: '15mm')
  filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
  return "(already exists, skipping) #{filename}" if File.exist?(filename)
  # while File.exist?(filename)
  #   counter = filename.match(/_(\d+)\.pdf/)[1] rescue 0
  #   counter += 1
  #   filename = filename.sub(/(?:_\d+)?\.pdf/, "_#{counter}.pdf")
  # end
  kit.to_file(filename)
  return "(saved) "+filename
end

def handler_pdf_from_html(message)
  body = message.html_part.decoded
  if body.match(/cid:/)
    # embedded files
    message.parts.select(&:content_id).each do |part|
      cid = part.content_id[1..-2]
      body.gsub!(/cid:#{cid}/, "data:#{part.content_type};base64,#{Base64::strict_encode64(part.decoded)}")
    end
  end
  kit = PDFKit.new(TEXT_EMAIL_BEFORE + helper_mail_header_html(message) + body + TEXT_EMAIL_AFTER, page_size: 'A4', margin_top: '15mm', margin_left: '15mm', margin_right: '15mm', margin_bottom: '15mm')
  filename = "#{message.date.strftime('%Y%m%dT%H%M%S')}__#{message.header[:from].decoded.gsub(/[^a-z0-9]/i, '_')}__#{message.message_id.split('@').first.gsub(/[^a-z0-9]/i, '_')}.pdf"
  return "(already exists, skipping) #{filename}" if File.exist?(filename)
  # while File.exist?(filename)
  #   counter = filename.match(/_(\d+)\.pdf/)[1] rescue 0
  #   counter += 1
  #   filename = filename.sub(/(?:_\d+)?\.pdf/, "_#{counter}.pdf")
  # end
  kit.to_file(filename)
  return "(saved) "+filename
end

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
    if File.exist?(filename)
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
    if File.exist?(filename)
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

SERVERS.each do |server|
  get_messages server[:host], server[:port], server[:ssl], server[:username], server[:password]
end
