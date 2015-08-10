# XlsReporter

XlsReporter lets you to export a collection of data to an Excel file, in a **Ruby** or **Ruby on Rails** project. Only need to specify the way to show data in the report.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'xls_reporter'
```
And then execute:
```sh
$ bundle install
```
Or install it yourself as:
```sh
$ gem install xls_reporter
```

## Usage
To export data:
```ruby
# create instance of reporter class
table_reporter = XlsReporter::DataExport.new

# generate report file
table_reporter.generate_excel(collection, header_params, body_params,
                          filename, filepath, batch_iteration, retry_methods)
```
with the follow parameters:
- *collection*: set of iterable data (`ActiveRecord::Relation`, `Array` of objects, etc.).
- *header_params*: `Array` of hashes, indicate the title of each column:
```ruby
  header_params = [
                {:index => 0, :title => "SKU"},
                {:index => 1, :title => "Event"},
                {:index => 2, :title => I18n.t(model.attribute)},
        ]
```
- *body_params*: `Array` of hashes, indicate the methods (and their params) to apply in each element of collection to show in report:
```ruby
  body_params = [
             # one method
             {:index => 0, :object_methods => :get_sku },
             # one method with one param
             {:index => 1, :object_methods => {:get_event => Time.now } },
             # method chaining with params
             {:index => 2, :object_methods => [ 
                  # one method with multiple params
                  {:method_name=>:values_by_ids, :params_values => [1,2,4]}, 
                  # one method
                  :get_color , 
                  #one method with one param
                  {:method_name =>:set_size, :params_values => "large"}
                ] 
             },
          ]
```
- *filename*: `String` with the name of the exported file (ie: *report.xlsx*).
- *filepath*: `String` with the path for save the report file.
- *batch_iteration*: this `boolean` option permit iterate data in [batches](http://api.rubyonrails.org/classes/ActiveRecord/Batches.html#method-i-find_each) if you use `ActiveRecord`. Default is **true**.
- *retry_methods*:  this `boolean` option is used to manage exceptions when calling methods in data elements, and retry them. Default is **true**.

## TO DO
  Write tests and validate code. 

## Contributing

1. Fork it ( https://github.com/[my-github-username]/xls_reporter/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
