require 'spreadsheet'
class CustomXlsWriter

	attr_accessor :fields, :collection, :path
	attr_accessor :book, :sheet

	attr_accessor :sheet_name, :column_headers

	#Fields: Array de hashes con la lista de campos que se incluiran:
	# => name: string con el nombre del campo
	# => value: string con el método para obtener su valor
	# => arguments: array<string> con los parámetros que recibe el método <value>
	# => type: string con el tipo del campo (string, currency, date)
	#collection: ActiveRecord_relation o array con la lista de objetos a exportar al xls
	#Path: string con el path en disco donde almacenar el archivo
	#options: Hash con opciones de personalización:
	# => sheet_name: string con el nombre de la hoja
	def initialize(_fields, _collection, _path, options = {})
		self.fields     = _fields
		self.collection = _collection
		self.path 		= _path

		options[:sheet_name].blank? ? self.sheet_name = "Sheet1" : self.sheet_name = options[:sheet_name]
		options[:column_headers].blank? ? self.column_headers = [] : self.column_headers = options[:column_headers]

		write()
	end
	
	def write
	    Spreadsheet.client_encoding = 'UTF-8'
	    self.book  = Spreadsheet::Workbook.new
	    self.sheet = book.create_worksheet :name => self.sheet_name

	    write_header()
	    write_objs()

		_check_path()
	    self.book.write self.path

	    return true
	end

private

	def write_header
	    header = self.sheet.row(0)
		if column_headers and column_headers.size > 0
			header.push ""
		end

	    for field in self.fields
	      header.push field[:name]
	    end
	end

	def write_objs
	    i = 1
	    for obj in self.collection
	      row = sheet.row(i)

		  if column_headers and column_headers.size > 0
		  	row.push column_headers[i-1]
		  end

	      for field in self.fields
	      	if field[:arguments] and !field[:arguments].blank?
	      		_value = obj.send(field[:value], *field[:arguments])
	      	else
	      		_value = obj.send(field[:value])
	      	end

	        case field[:type] 
	          when "string"
	            row.push _value
	          when "currency"
	            row.push ActionController::Base.helpers.number_to_currency(_value)
	          when "date"
	          	val = I18n.l(_value) rescue ''
	            row.push(val)
	        end
	      end

	      i += 1
	    end
	end

private

	def _check_path
		require 'fileutils'

		aux_path = self.path.split("/")[0...-1].join("/")
		unless File.directory?(aux_path)
			FileUtils::mkdir_p aux_path
		end
	end

end
