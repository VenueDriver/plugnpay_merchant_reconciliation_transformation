require 'roo'
require 'rubyXL'

# Get input file name.
unless filename = ARGV[0] and File.exist?(filename)
  puts "Usage: #{$0} spreadsheet.xlsx"
  exit
end

# Open the spreadsheet and get basic info.
xlsx = Roo::Spreadsheet.open(filename)
puts xlsx.info

# Locate PlugNPay sheet by finding the first sheet with "VisaMC" in the name.
plugnpay_sheet_name = xlsx.sheets.select{|sheet| sheet =~ /VisaMC/i}.first
puts "PlugNPay sheet: #{plugnpay_sheet_name.inspect}"

# Locate Venue Driver report sheet, also by looking for the name.
venuedriver_sheet_name = xlsx.sheets.select{|sheet| sheet =~ /VD\s*Report/i}.first
puts "Venue Driver report sheet: #{venuedriver_sheet_name.inspect}"

# Derive the output file name from the input file name.
outfile = filename.gsub(/(\.\w+$)/,'-transformed\1')
puts "Output file: #{outfile}"

# Build an in-memory hash of ticket sale rows from the Venue Driver data,
# by order ID.
ticket_sales_by_order_id = {}
(venuedriver_sheet = xlsx.sheet(venuedriver_sheet_name)).
  each_with_index(orderID: 'order_id') do |hash, index|
  # For each row, get a hash with the order ID, plus an array of raw cells.
  row_number = index + 1
  row = venuedriver_sheet.row(row_number)

  # Create an array for ticket sales for this order ID if necessary.
  ticket_sales_by_order_id[hash[:orderID]] ||= []
  # Add this ticket sale row to the bucket for this order ID.
  ticket_sales_by_order_id[hash[:orderID]] << row
end
venue_driver_header_slice = venuedriver_sheet.row(1).slice(1,15)

# Loop through each row in the PlugNPay sheet.
output_rows = []
(plugnpay_sheet = xlsx.sheet(plugnpay_sheet_name)).
  each_with_index(orderID: 'PnP OrderID') do |hash, index|
  # For each row, get a hash with the order ID, plus an array of raw cells.
  row_number = index + 1
  row = plugnpay_sheet.row(row_number)

  puts "row: #{row_number}"
  puts "raw: #{row}"
  puts hash.inspect

  # Find the tidket sales for this order ID from the Venue Driver sheet.
  order_id_ticket_sales = ticket_sales_by_order_id[hash[:orderID]] || []

  # For the header row.
  if index == 0
    output_rows << [
      row.slice(0,4),
      venue_driver_header_slice,
      row.slice(4,row.length)
    ].flatten
  # If this row has an order ID, and...
elsif order_id_ticket_sales.length > 0
    # Emit an output row for each
    order_id_ticket_sales.each do |raw_sale_row|
      sliced_sale_row = raw_sale_row.slice(1,15)

      output_rows << [
        row.slice(0,4),
        sliced_sale_row,
        row.slice(4,row.length)
      ].flatten
    end
  else
    output_rows << [
      row.slice(0,4),
      Array.new(15, nil),
      row.slice(4,row.length)
    ].flatten
  end

  puts "output: " + output_rows.last.inspect
  puts '-----'
end

# Write the output_rows array to a file.
workbook = RubyXL::Workbook.new
worksheet = workbook[0]
output_rows.each_with_index do |row, row_number|
  row.each_with_index do |cell, column_number|
    worksheet.add_cell(row_number, column_number, cell)
  end
end
workbook.write(outfile)
