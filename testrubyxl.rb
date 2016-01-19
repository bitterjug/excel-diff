require 'optparse'
require 'ostruct'
require 'rubyXL'
require 'pry'

usage = "Usage: ruby #{__FILE__} [options] <workbook-file>"

$options = OpenStruct.new
# defaults:
$options.values = true
$options.formulas = true

OptionParser.new do |opts|
  opts.banner = usage

  opts.on("-v", "--no-values", "Exclude cell values") do |v|
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
    ref = RubyXL::Reference.ind2ref(cell.column, cell.row) # TODO should this be the other way round?
    sheet = cell.worksheet.sheet_name  #  TODO normalize and transcode
    "#{sheet}!#{ref}"
end

#
# TODO: This should probably behave differently if we're hiding calculated values
# then we should show all literal values, and for calculated values show some hint 
# that its a formula. We cant lways showt he formula because of shred formula ranges
#
def value(cell)
    $options.values ? " = #{cell.value}" : ": "
end

def formula(cell)
    if $options.formulas and cell.formula
        range = cell.formula.t == "shared" ?  "[#{cell.formula.ref.to_s}]" : ""
        return "\t#{cell.formula.t} formula #{range} =#{cell.formula.expression}"
    else
        return ""
    end
    # cell.formula.ref is a Referenec object !! test with cross sheet links?
end

# TODO:  Unicode normalize NFKD and convert to ascii? to make diff work
# https://github.com/knu/ruby-unf
def ouput(cell)
    puts "#{address(cell)}#{value(cell)}"
    puts formula(cell)
end

def process(wb)
    wb.each do |ws|
        ws.each do |row|
            row && row.cells.each do |cell|
                ouput(cell)
                # binding.pry
            end
            puts
        end
    end
end


filename = File.expand_path(ARGV[0])
wb = RubyXL::Parser.parse(filename)
process(wb)
