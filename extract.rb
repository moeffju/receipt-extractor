#!/usr/bin/env ruby
# encoding: utf-8

require 'rubygems'
require 'bundler/setup'

require 'pdf-reader'
require 'pry'

# this extracts amounts from PDFs
if ARGV.first == '--help' || ARGV.first == '-h'
  STDERR.puts <<_EOF_
Extracts invoice amounts from PDF invoices and generates tab-separated output
listing supplier, invoice date and invoice amount, suitable for pasting into the
mobility package Excel sheet.

You probably want to redirect the output of this script to a file.

Syntax:

    $0 [pdf-file] [pdf-file...] >output.tsv

Arguments:

    pdf file…
        Pass any PDF files (as generated by `fetch.rb`) that you want to parse
_EOF_
  exit 1
end

def refmt_coup_date(str)
  "#{str[-7..-6]}.#{str[0..1]}.#{str[-4..-1]}"
end

ARGV.each do |filename|
  STDERR.puts "File not found: #{filename}!" and next unless File.exist?(filename)
  reader = PDF::Reader.new(filename)
  supplier = invoice_date = invoice_amount = ''
  full_text = reader.pages.map(&:text).join("\n--PAGE_BREAK--\n")
  if full_text =~ /mytaxi ID:/
    supplier = 'mytaxi Intelligent Apps GmbH'
    invoice_date = full_text.match(/Belegdatum.*:\s*(\d+\.\d+\.\d+) \d+:\d+$/u)[1]
    invoice_amount = full_text.match(/Bruttobetrag.*\s+(\d+,\d+)\s+€$/u)[1]
  elsif full_text =~ /BVG-OnlineShop/i
    supplier = "Berliner Verkehrsbetriebe"
    invoice_date = full_text.match(/Date:\ (?<invoice_date>[\d-]+)/u)[:invoice_date]
    invoice_amount = full_text.match(/Order\ total:\ €(?<invoice_amount>\d+\.\d+)/u)[:invoice_amount]
  elsif full_text =~ /(Electric Mobility Concepts GmbH|drive by mobility GmbH)/
    # this should work for all providers using fleetbird
    supplier = $1
    invoice_date = full_text.match(/Rechnungsdatum:\s+(?<invoice_date>[\d.]+)/u)[:invoice_date]
    invoice_amount = full_text.match(/Gesamtbetrag:\s+(?<invoice_amount>\d+\.\d+) EUR/u)[:invoice_amount]
  elsif full_text =~ /(car2go Deutschland GmbH)/
    supplier = $1
    invoice_date = full_text.match(/(Rechnungsdatum|Datum):\s+(?<invoice_date>[\d.]+)/u)[:invoice_date]
    invoice_amount = full_text.match(/Gesamtbetrag\s+(?:\d+,\d+\s+){2}(?<invoice_amount>\d+\,\d+)/u)[:invoice_amount]
  elsif full_text =~ /(DriveNow GmbH \& Co\. KG)/
    supplier = $1
    invoice_date = full_text.match(/Berlin, (?<invoice_date>[\d.]+)/u)[:invoice_date]
    invoice_amount = full_text.match(/Gesamtkosten:\s+(?<invoice_amount>\d+\,\d+) EUR/u)[:invoice_amount]
  elsif full_text =~ /(Coup Mobility GmbH)/
    supplier = $1
    invoice_date = refmt_coup_date(full_text.match(/Invoice\sdate:\s+(?<invoice_date>[\d\-]+)/u)[:invoice_date])
    invoice_amount = full_text.match(/Total\samount\s+(?<invoice_amount>\d+\.\d+) €/u)[:invoice_amount]
  elsif full_text =~ /From:\s*(Uber) Receipts/
    if full_text =~ /Personal/
      supplier = 'Uber'
      invoice_date = full_text.match(/Date:\s+(?<invoice_date>[\d\-]+)/u)[:invoice_date]
      invoice_amount = full_text.scan(/\d+[.,]\d+/u).first
    end
  elsif full_text =~ /bahncard.service@bahn.de/
    supplier = 'DB Fernverkehr'
    invoice_date = full_text.match(/(?<invoice_date>\d+\.\d+\.\d{4})/u)[:invoice_date]
    invoice_amount = full_text.match(/Gesamtbetrag\s+(?<invoice_amount>\d+,\d+)\s+€/u)[:invoice_amount]
  elsif full_text =~ /Online-Ticket/ and full_text =~ /Fernverkehr/
    supplier = 'DB Fernverkehr'
    invoice_date = full_text.match(/erfolgte am (?<invoice_date>\d+\.\d+\.\d{4})/u)[:invoice_date]
    invoice_amount = full_text.match(/Summe\s+(?<invoice_amount>\d+,\d+)€/u)[:invoice_amount]
  elsif full_text =~ /Online seat reservation/ and full_text =~ /bahn\.com/
    supplier = 'DB Fernverkehr'
    invoice_date = full_text.match(/reservation was made on (?<invoice_date>\d+\.\d+\.\d{4})/u)[:invoice_date]
    invoice_amount = full_text.match(/Total fare\s+(?<invoice_amount>\d+,\d+)\s*€/u)[:invoice_amount]
  elsif full_text =~ /(Deutsche Bahn Connect GmbH)/
    supplier = $1
    invoice_date = full_text.match(/, den (?<invoice_date>\d+\.\d+\.\d{4})/u)[:invoice_date]
    invoice_amount = full_text.match(/Gesamtbetrag\s+(?<invoice_amount>\d+,\d+)\s*€/u)[:invoice_amount]
  else
    STDERR.puts full_text
    binding.pry
  end
  
  invoice_amount.gsub!(/\./, ',')
  #puts "#{filename}: #{supplier} - #{invoice_amount} (#{invoice_date})"
  puts [supplier, invoice_amount, invoice_date].join("\t")
end

