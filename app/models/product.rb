class Product < ApplicationRecord

    require 'roo'

    def self.import(file)
        spreadsheet = Roo::Spreadsheet.open(file.path)
        header = spreadsheet.row(1)
        (2..spreadsheet.last_row).each do |i|
        row = Hash[[header, spreadsheet.row(i)].transpose]        
        product = find_by(id: row["id"]) || new
        product.attributes = row.to_hash
        product.save!
        end
    end
end
