require "xls_reporter/version"

module XlsReporter
  class DataExport
      require 'write_xlsx'

      NUMBER_OF_RETRIES = 5

      # generate_excel: Main method to generate file
      # ----------------------------------------------------------------------------
      # Valid header_params format:
      #    header_params = [
      #       {:index => 0, :title => I18n.t(model.attributes)},
      #       {:index => 1, :title => "Event"},
      #       {:index => 2, :title => "SKU"},
      #    ]
      # ----------------------------------------------------------------------------
      # Valid body_params format:
      #    body_params = [
      #       {:index => 0, :object_methods => [ {:method_name=>:values_by_ids, :params_values => [1,2,4]}, :get_color , {:method_name =>:set_size, :params_values => "large"}] },
      #       {:index => 1, :object_methods => {:get_event => Time.now } },
      #       {:index => 2, :object_methods => :get_sku },
      #    ]
      def generate_excel collection, header_params, body_params, filename="report_#{Time.now.strftime("%Y%m%d%H%M")}.xlsx", filepath = default_path, batch_iteration = true, retry_methods = true

         fullpath= File.join(filepath, filename)
         puts "#{Time.now} Generating #{fullpath}"

         workbook = WriteXLSX.new(fullpath)
         worksheet = workbook.add_worksheet
         row = 0

         # Set Headers
         if header_params
            header_params.each do |param|
               worksheet.write(row, param[:index], param[:title])
            end
            row = row + 1
         end

         # Use batch iteration if ActiveRecord is defined and param is set
         iterate_method = defined?(ActiveRecord) && collection.is_a?(ActiveRecord::Relation)  && batch_iteration ? "find_each" : "each"

         if body_params
            collection.send(iterate_method) do |object|
               body_params.each do |param|

                  # Get object column value depending of datatype of args
                  object_value = if param[:object_methods].is_a? Array #Multiple methods
                     #Chaining methods across the array
                     param[:object_methods].reduce(object) do |method_params, chained_value|
                        call_method chained_value, method_params, retry_methods ? 0 : -1
                     end
                  else
                    call_method object, param[:object_methods], retry_methods ? 0 : -1
                  end

                  worksheet.write(row, param[:index], object_value)
               end

               row = row + 1
            end
         end

         workbook.close
         puts "#{Time.now} Exporting to #{fullpath} done!"

       end


      private

      # Call methods for each object and retry if fail
        def call_method object, method_params, retry_count = -1
           begin
              if method_params.is_a? Hash
                if method_params.key?(:method_name) && method_params.key?(:params_values) #  {:method _name=>:values_by_ids, :params_values => [1,2,4]}
                  method_name = method_params[:method_name]
                  object.send method_name, *method_params[:params_values]
                else #  {:get_event => Time.now }
                  method_name = method_params.flatten.first
                  object.send method_name.to_sym, method_params[method_name]
                end
              else # method_name
                method_name = method_params
                object.send method_name.to_sym
              end
           rescue Exception => e
              puts "Error calling method  method_name. #{e}"

              if retry_count > 0 && retry_count <= NUMBER_OF_RETRIES
                seconds_to_sleep = (Random.rand(5) + 5)*(retry_count+1) # (5 + 0..5) x (#retries + 1) seconds
                puts "Retry call method in #{seconds_to_sleep} seconds"

                sleep seconds_to_sleep
                call_method object, method_params, retry_count+1
              else
                puts "Retries finished or ignored, filling with blank"
                ""
              end
           end
        end

        # Rails defined?
        def rails?
          defined? Rails
        end

        #Default values of  main method
        def default_path
           rails? ? Rails.public_path : File.expand_path(File.dirname(__FILE__))
        end

   end
end
