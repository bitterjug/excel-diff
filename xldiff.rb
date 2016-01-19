require 'optparse'
require 'ostruct'
require 'rubyXL'

usage = "Usage: ruby #{__FILE__} [options] <workbook-file>"

$options = OpenStruct.new
# defaults:
$options.values = true
$options.formulas = true

OptionParser.new do |opts|
  opts.banner = usage

  opts.on("-v", "--no-values", "Exclude calculated values") do |v|
    $options.values = v
    # TODO: should this be for calclated values only?
  end

  opts.on("-f", "--no-formulas", "Exclude cell formulas") do |v|
    $options.formulas = v
  end

  opts.on_tail("-h", "--help", "Show this message") do
    puts opts
    puts "help"
    exit
  end

end.parse!

if ARGV.length != 1
    puts usage
    exit
end

def address(cell)
    ref = RubyXL::Reference.ind2ref(cell.row, cell.column) 
    sheet = cell.worksheet.sheet_name  #  TODO normalize and transcode
    "#{sheet}!#{ref}"
end

def value(cell)
    (not $options.values and cell.formula) ? ": " : " = #{cell.value}"
end

def formula(cell)
    if $options.formulas and cell.formula
        range = cell.formula.t ?  "[#{cell.formula.ref.to_s}] " : ""
        type = cell.formula.t ? "#{cell.formula.t} " : ""
        return "#{$options.values ? "\n\t" : ""}#{type}formula #{range}=#{cell.formula.expression}"
    else
        return ""
    end
end

# TODO:  Unicode normalize NFKD and convert to ascii? to make diff work
# https://github.com/knu/ruby-unf
def ouput(cell)
    puts "#{address(cell)}#{value(cell)}#{formula(cell)}"
end

def process(wb)
    wb.each do |ws|
        ws.each do |row|
            row && row.cells.each do |cell|
                ouput(cell)
            end
            puts
        end
    end
end


filename = File.expand_path(ARGV[0])
wb = RubyXL::Parser.parse(filename)
process(wb)
