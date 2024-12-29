require "google/cloud/logging"
require "time"
require 'active_support/core_ext/numeric/time'
require 'google/apis/sheets_v4'
require 'googleauth'
require "time"

class GetLogBugService
  def execute
    project_id = "knoock-core"

    sheets_service = service_google_sheet
    spreadsheet_id = '1i3tMD25dqtV7rHyakCuixtujF9uiMwrkBvxuHw24-kE'

    prepare_sheet(spreadsheet_id)

    title_range = "#{sheet_name}!13:13"
    title_response = sheets_service.get_spreadsheet_values(spreadsheet_id, title_range)
    titles = title_response.values.first

    date_col_index = titles.index('Date (JST)')
    detail_col_index = titles.index('Error Message')
    count_col_index = titles.index('Number')
    url_loggin_col_index = titles.index('URL')


    logging = Google::Cloud::Logging.new(project: project_id)

    current_time = Time.now.in_time_zone("Asia/Tokyo")
    eight_am = current_time.beginning_of_day + 8.hours
    start_at = (eight_am - 1.days).beginning_of_day.utc.iso8601
    end_at = current_time.end_of_day.utc.iso8601

    entries_mysql = logging.entries(filter: prepare_filter_search("Mysql2::Error", start_at, end_at))
    entries_lost_connect = logging.entries(filter: prepare_filter_search("Lost connection", start_at, end_at))
    entries_ruby = logging.entries(filter: prepare_filter_search("/app/vendor/bundle/ruby/", start_at, end_at))
    entries_error = logging.entries(filter: prepare_filter_search("error", start_at, end_at))
    entries_severity_error = logging.entries(filter: filter_severity_error(start_at, end_at))

    entries = entries_mysql + entries_lost_connect + entries_ruby + entries_error + entries_severity_error
    entries = entries.flatten
    entries = entries.sort_by {|e| e.timestamp}

    errors = Hash.new { |hash, key| hash[key] = { numbers: 0, time: nil, url: nil, is_duplicate: false } }
    entries.each do |entry|
      if entry.payload.is_a?(Google::Protobuf::Struct)
        next unless  entry.payload.fields["message"]
        message = entry.payload.fields["message"].string_value
      else
        message = entry.payload.inspect
      end

      timestamp = entry.timestamp.in_time_zone("Asia/Tokyo")
      normalized_message = normalize_message(message)

      next unless normalized_message

      if errors.key?(normalized_message)
        errors[normalized_message][:numbers] += 1
      else
        errors[normalized_message] = {
          numbers: 1,
          time: timestamp,
          url: "https://console.cloud.google.com/logs/query;startTime=#{(entry.timestamp - 2.seconds).utc.iso8601};endTime=#{entry.timestamp.utc.iso8601}?referrer=search&project=#{project_id}"
        }
      end
    end

    values = errors.uniq.map do |message, data|
      safe_message = message.to_s.encode('UTF-8', invalid: :replace, undef: :replace, replace: '?')
      date = data[:time]

      row = Array.new(titles.size)
      row[date_col_index] = date
      row[detail_col_index] = safe_message
      row[count_col_index] = data[:numbers]
      row[url_loggin_col_index] = data[:url]
      row
    end

    values.compact!
    errors.each do |message, infomation|
      puts "Log Message: #{message}"
      puts "Count: #{infomation[:numbers]}"
      puts "Time: #{infomation[:time]}"
      puts "----------------------------------"
    end

    sheet_row_number = sheets_service.get_spreadsheet_values(spreadsheet_id, "#{sheet_name}!A:Z").values.size

    range = "#{sheet_name}!A#{sheet_row_number}"
    value_range = Google::Apis::SheetsV4::ValueRange.new(values: values)

    sheets_service.append_spreadsheet_value(spreadsheet_id, range, value_range, value_input_option: 'RAW')

    sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)
    requests = []

    requests << {
      update_cells: {
        range: {
          sheet_id: sheet_id,
          start_row_index: sheet_row_number.to_i,
          end_row_index: sheet_row_number.to_i + values.size,
          start_column_index: 0,
          end_column_index: values.first.size
        },
        rows: values.each_with_index.map do |row, i|
          {
            values: row.map { |value| { user_entered_value: { string_value: value.to_s } } }
          }
        end,
        fields: "user_entered_value"
      }
    }

    response = sheets_service.get_spreadsheet_values(spreadsheet_id, "#{sheet_name}!A:Z")
    row_number = response.values.size

    (0..row_number).each do |row_index|
      (0..7).each do |col_index|
        requests << {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 13,
              end_row_index: 13 + row_index,
              start_column_index: col_index,
              end_column_index: col_index + 1
            },
            top: { style: 'SOLID', width: 1 },
            bottom: { style: 'SOLID', width: 1 },
            left: { style: 'SOLID', width: 1 },
            right: { style: 'SOLID', width: 1 }
          }
        }
      end
    end

    requests << {
      update_cells: {
        range: {
          sheet_id: sheet_id,
          start_row_index: 13,
          end_row_index: row_number,
          start_column_index: 4,
          end_column_index: 5
        },
        rows: values.each_with_index.map do |row, i|
          {
            values: [{
              user_entered_value: { string_value: row[4].to_s },
              user_entered_format: { wrap_strategy: 'CLIP' }
            }]
          }
        end,
        fields: "user_entered_value,user_entered_format.wrap_strategy"
      }
    }

    requests << {
      update_dimension_properties: {
        range: {
          sheet_id: sheet_id,
          dimension: 'COLUMNS',
          start_index: 2,
          end_index: 3
        },
        properties: {
          pixel_size: 800
        },
        fields: 'pixel_size'
      }
    }

    requests << {
      update_dimension_properties: {
        range: {
          sheet_id: sheet_id,
          dimension: 'COLUMNS',
          start_index: 1,
          end_index: 2
        },
        properties: {
          pixel_size: 200
        },
        fields: 'pixel_size'
      }
    }

    sheets_service.batch_update_spreadsheet(spreadsheet_id, { requests: requests })

    sum_bug_numbers(sheet_name, sheets_service, spreadsheet_id)

    puts "Dữ liệu đã được ghi vào Google Sheets thành công."
  end

  def normalize_message(message)
    if message.include?("Aborted connection") || message.include?("Got an error reading communication packets")
      "Aborted connection to db: 'knoock_api_production'"
    elsif message.include?("Error watching metadata")
      "Error watching metadata"
    elsif message.include?("NaN")
      "Error NaN"
    elsif message.include?('at settle (file:///app/node_modules/axios/lib/core/settle.js:19:12)') || message.include?('clarifyTimeoutError: false') ||
      message.include?("error: [Function (anonymous)],") || message.include?("authorizationError: null,") ||  message.include?("_hadError: false,") ||
      message.include?("_closeAfterHandlingError: false,") || message.include?("data: { error: 'Record not found' }")
      "AxiosError: Request failed with status code 404 at settle (file:///app/node_modules/axios/lib/core/settle.js:19:12) at IncomingMessage.handleStreamEnd (file:///app/node_modules/axios/lib/adapters/http.js:548:11) at IncomingMessage.emit (node:events:525:35) at endReadableNT (node:internal/streams/readable:1359:12) at process.processTicksAndRejections (node:internal/process/task_queues:82:21) {"
    elsif message.include?('ActiveRecord::LockWaitTimeout: Mysql2::Error::TimeoutError: Lock wait timeout exceeded; try restarting transaction') || message.include?('Mysql2::Error::TimeoutError: Lock wait timeout exceeded; try restarting transaction')
      'ActiveRecord::LockWaitTimeout: Mysql2::Error::TimeoutError: Lock wait timeout exceeded; try restarting transaction'
    elsif message.include?('gems/mysql2-0.5.4/lib/mysql2/client.rb:148') || message.include?('gems/mysql2-0.5.4/lib/mysql2/client.rb:147') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_mysql_adapter.rb:632') || message.include?('gems/activesupport-7.0.4/lib/active_support/concurrency/share_lock.rb:187') ||
      message.include?('gems/activesupport-7.0.4/lib/active_support/dependencies/interlock.rb:41') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_mysql_adapter.rb:631') ||
      message.include?('gems/activesupport-7.0.4/lib/active_support/concurrency/load_interlock_aware_monitor.rb:25') || message.include?('gems/activesupport-7.0.4/lib/active_support/concurrency/load_interlock_aware_monitor.rb:21') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_adapter.rb:765') || message.include?('gems/activesupport-7.0.4/lib/active_support/notifications/instrumenter.rb:24') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_mysql_adapter.rb:630') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/mysql/database_statements.rb:96:') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/mysql/database_statements.rb:47') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_mysql_adapter.rb:207') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/mysql/database_statements.rb:52') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/database_statements.rb:560') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/database_statements.rb:66') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/query_cache.rb:110') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/mysql/database_statements.rb:12') || message.include?('gems/activerecord-7.0.4/lib/active_record/querying.rb:54') || message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:942') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:962') || message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:928') || message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:914') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:908') || message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:695') || message.include?('gems/activerecord-7.0.4/lib/active_record/relation.rb:250') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/relation/delegation.rb:88') || message.include?('gems/activesupport-7.0.4/lib/active_support/core_ext/enumerable.rb:106') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/transaction.rb:319') ||
      message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/transaction.rb:317') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract/database_statements.rb:316') || message.include?('gems/activerecord-7.0.4/lib/active_record/transactions.rb:209') ||
      message.include?('gems/rake-13.0.6/lib/rake/task.rb:281') || message.include?('gems/rake-13.0.6/lib/rake/task.rb:219') ||
      message.include?('gems/rake-13.0.6/lib/rake/task.rb:199') || message.include?('gems/rake-13.0.6/lib/rake/task.rb:188') || message.include?('gems/activerecord-7.0.4/lib/active_record/connection_adapters/abstract_adapter.rb:756') || message.include?('""') ||
      message.include?('Start GoogleIndexNotifyContentUpdateService') || message.include?('End GoogleIndexNotifyContentUpdateService') || message.include?('GoogleIndexNotifyContentUpdateService::IndexingApiRequestError')
      nil
    elsif message.include?('/rake-13.0.6/lib/rake/application.rb:160') || message.include?('/rake-13.0.6/lib/rake/application.rb:116') || message.include?('/rake-13.0.6/lib/rake/application.rb:125') || message.include?('rake-13.0.6/lib/rake/application.rb:110') ||
      message.include?('railties-7.0.4/lib/rails/commands/rake/rake_command.rb:24') || message.include?('railties-7.0.4/lib/rails/commands/rake/rake_command.rb:24') || message.include?('rake-13.0.6/lib/rake/application.rb:186') ||
      message.include?('rake-13.0.6/lib/rake/rake_module.rb:59') || message.include?('ems/railties-7.0.4/lib/rails/commands/rake/rake_command.rb:18') || message.include?('gems/railties-7.0.4/lib/rails/command.rb:51') ||
      message.include?('app/vendor/bundle/ruby/3.1.0/gems/bootsnap-1.15.0/lib/bootsnap/load_path_cache/core_ext/kernel_require.rb:32')
      '"    \"message\": \"Failed to parse batch request, error:  0 items. Received batch body: (49) bytes redacted\","'
    elsif message.include?('Google::Protobuf::Any: type_url: "type.googleapis.com/google.cloud.audit.AuditLog')
      'Google::Protobuf::Any: type_url: "type.googleapis.com/google.cloud.audit.AuditLog'
    elsif message.include?('Out of sort memory, consider increasing server sort buffer size')
      'Mysql2::Error: Out of sort memory, consider increasing server sort buffer size'
    else
      message
    end
  end

  def prepare_filter_search(content, start_at, end_at)
    filter = <<-FILTER
      (textPayload:"#{content}" OR
      jsonPayload.message:"#{content}") AND
      timestamp >= "#{start_at}" AND
      timestamp <= "#{end_at}" AND
      NOT (textPayload:"gce_workload_cert_refresh" OR jsonPayload.message:"gce_workload_cert_refresh")
    FILTER
    filter
  end

  def filter_severity_error(start_at, end_at)
    filter = <<-FILTER
      severity="ERROR" AND
      timestamp >= "#{start_at}" AND
      timestamp <= "#{end_at}"
    FILTER
    filter
  end

  def date_of_week
    today = Date.today - 1.days
    {
      start_of_week: today.beginning_of_week(:monday),
      end_of_week: today.end_of_week(:monday)
    }
  end

  def sheet_name
    today = Date.today - 1.days
    start_of_week = date_of_week[:start_of_week]
    end_of_week = date_of_week[:end_of_week]
    "#{start_of_week.strftime('%d')}-#{end_of_week.strftime('%d')}/#{start_of_week.strftime('%m')}/#{start_of_week.strftime('%Y')}"
  end

  def prepare_sheet(spreadsheet_id)
    project_id = "knoock-core"
    sheets_service = service_google_sheet

    create_sheet_if_not_exists(sheets_service, spreadsheet_id, sheet_name)
  end

  def create_sheet_if_not_exists(sheets_service, spreadsheet_id, sheet_name)
    spreadsheet = sheets_service.get_spreadsheet(spreadsheet_id)
    existing_sheets = spreadsheet.sheets.map { |sheet| sheet.properties.title }

    unless existing_sheets.include?(sheet_name)
      new_sheet_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
        requests: [
          add_sheet: { properties: { title: sheet_name } }
        ]
      )
      sheets_service.batch_update_spreadsheet(spreadsheet_id, new_sheet_request)
      puts "Created new sheet: #{sheet_name}"

      start_of_week = date_of_week[:start_of_week]
      end_of_week = date_of_week[:end_of_week]
      sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)
      requests = [
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 0,
              end_row_index: 1,
              start_column_index: 0,
              end_column_index: 1
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "ERROR SUMMARY" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 12 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 1,
              end_row_index: 2,
              start_column_index: 1,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "FROM" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 10 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 1,
              end_row_index: 2,
              start_column_index: 2,
              end_column_index: 3
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: start_of_week },
                    user_entered_format: {
                      text_format: { font_size: 10 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 2,
              end_row_index: 3,
              start_column_index: 1,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "TO" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 10 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 2,
              end_row_index: 3,
              start_column_index: 2,
              end_column_index: 3
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: end_of_week },
                    user_entered_format: {
                      text_format: { font_size: 10 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 4,
              end_row_index: 5,
              start_column_index: 0,
              end_column_index: 1
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "SUMMARY" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 12 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 5,
              end_row_index: 6,
              start_column_index: 1,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Last week" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 10 }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat.textFormat"
          }
        },
        {
          merge_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6,
              end_row_index: 7,
              start_column_index: 0,
              end_column_index: 2
            },
            merge_type: "MERGE_ALL"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6,
              end_row_index: 7,
              start_column_index: 0,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_format: {
                      background_color: {
                        red: 0.9,
                        green: 0.9,
                        blue: 0.9
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredFormat.backgroundColor"
          }
        },
        # C7: đặt background là light gray 1, text NUMBER in đậm, font size 12
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6,
              end_row_index: 7,
              start_column_index: 2, # Cột C
              end_column_index: 3
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "NUMBER" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 11 },
                      background_color: {
                        red: 0.9,
                        green: 0.9,
                        blue: 0.9
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6,
              end_row_index: 7,
              start_column_index: 3,
              end_column_index: 4
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "NOTE" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 11 },
                      background_color: {
                        red: 0.9,
                        green: 0.9,
                        blue: 0.9
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },
        {
          merge_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 7,
              end_row_index: 8,
              start_column_index: 0,
              end_column_index: 2
            },
            merge_type: "MERGE_ALL"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 7,
              end_row_index: 8,
              start_column_index: 0,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Number of bugs this week" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 11 },
                      background_color: {
                        red: 0.9,
                        green: 0.9,
                        blue: 0.9
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },
        # Gộp A9 và B9, text "Compared to last week", in đậm, font size 12
        {
          merge_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 8,
              end_row_index: 9,
              start_column_index: 0,
              end_column_index: 2
            },
            merge_type: "MERGE_ALL"
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 8,
              end_row_index: 9,
              start_column_index: 0,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Compared to last week" },
                    user_entered_format: {
                      text_format: { bold: true, font_size: 11 },
                      background_color: {
                        red: 0.9,
                        green: 0.9,
                        blue: 0.9
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },
        {
          merge_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 11,
              end_row_index: 12,
              start_column_index: 0,
              end_column_index: 8
            },
            merge_type: 'MERGE_ALL'
          }
        },
        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 11,
              end_row_index: 12,
              start_column_index: 0,
              end_column_index: 8
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "ERROR DETAIL" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 12
                      },
                      horizontal_alignment: 'CENTER',
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 0,
              end_column_index: 1
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "No" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 1,
              end_column_index: 2
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Date (JST)" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 2,
              end_column_index: 3
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Error Message" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 3,
              end_column_index: 4
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Number" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 4,
              end_column_index: 5
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "URL" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 5,
              end_column_index: 6
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Report URL" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 6,
              end_column_index: 7
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Status" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },

        {
          update_cells: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 7,
              end_column_index: 8
            },
            rows: [
              {
                values: [
                  {
                    user_entered_value: { string_value: "Note" },
                    user_entered_format: {
                      text_format: {
                        bold: true,
                        font_size: 11
                      },
                      background_color: {
                        red: 1.0,
                        green: 0.8,
                        blue: 0.8
                      }
                    }
                  }
                ]
              }
            ],
            fields: "userEnteredValue,userEnteredFormat"
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6, # Row 7
              end_row_index: 7,
              start_column_index: 0, # Column A
              end_column_index: 2 # Column B
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        # C7: Border for C7
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6, # Row 7
              end_row_index: 7,
              start_column_index: 2, # Column C
              end_column_index: 3
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        # D7: Border for D7
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 6, # Row 7
              end_row_index: 7,
              start_column_index: 3, # Column D
              end_column_index: 4
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        # A8:B8: Border for A8 and B8
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 7, # Row 8
              end_row_index: 8,
              start_column_index: 0, # Column A
              end_column_index: 2 # Column B
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 7, # Row 8
              end_row_index: 8,
              start_column_index: 2, # Column D
              end_column_index: 3
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 7, # Row 8
              end_row_index: 8,
              start_column_index: 3, # Column D
              end_column_index: 4
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 8, # Row 9
              end_row_index: 9,
              start_column_index: 2, # Column D
              end_column_index: 3
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 8, # Row 9
              end_row_index: 9,
              start_column_index: 3, # Column D
              end_column_index: 4
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        # A9:B9: Border for A9 and B9
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 8, # Row 9
              end_row_index: 9,
              start_column_index: 0, # Column A
              end_column_index: 2 # Column B
            },
            top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } }
          }
        },
        {
          update_borders: {
            range: {
              sheet_id: sheet_id,
              start_row_index: 12,
              end_row_index: 13,
              start_column_index: 0,
              end_column_index: 8
            },
            borders: {
              top: { style: "SOLID", width: 1 },
              bottom: { style: "SOLID", width: 1 },
              left: { style: "SOLID", width: 1 },
              right: { style: "SOLID", width: 1 }
            }
          }
        }
      ]
      batch_update_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(requests: requests)
      sheets_service.batch_update_spreadsheet(spreadsheet_id, batch_update_request)
    end
  end

  def service_google_sheet
    sheets_service = Google::Apis::SheetsV4::SheetsService.new
    sheets_service.authorization = Google::Auth::ServiceAccountCredentials.make_creds(
      json_key_io: File.open('credentials.json'),
      scope: ['https://www.googleapis.com/auth/spreadsheets']
    )
    puts "OK credentials sheet"
    sheets_service
  end

  def get_sheet_id(sheets_service, spreadsheet_id, sheet_name)
    spreadsheet = sheets_service.get_spreadsheet(spreadsheet_id)
    sheet = spreadsheet.sheets.find { |s| s.properties.title == sheet_name }
    sheet.properties.sheet_id
  end

  def sum_bug_numbers(sheet_name, sheets_service, spreadsheet_id)
    sheet_row_number = sheets_service.get_spreadsheet_values(spreadsheet_id, "#{sheet_name}!A:Z").values.size + 1
    total = sheet_row_number - 14

    result_range = "#{sheet_name}!C8"
    value_range = Google::Apis::SheetsV4::ValueRange.new(values: [[total]])
    sheets_service.update_spreadsheet_value(
      spreadsheet_id,
      result_range,
      value_range,
      value_input_option: 'USER_ENTERED'
    )
  end
end

GetLogBugService.new().execute
