module Agents
  class OutlookAgent < Agent
    include FormConfigurable
    can_dry_run!
    no_bulk_receive!
    default_schedule '5m'

    description do
      <<-MD
      The Outlook Agent interacts with the graph api to check the arrival of new emails, send them, etc...

      `debug` is used for verbose mode.

      `client_id` is the id of your app.

      `client_secret` is the secret of your app.

      `access_token` is a token created for your app.

      `refresh_token` is needed to refresh access_token.

      `folder` is the wanted folder where you want to check new emails.

      `raw_email` if you want MIME format content.

      `emit_events` is for creating an event.

      `type` is for the wanted action like get_new_emails / send_email.

      `expected_receive_period_in_days` is used to determine if the Agent is working. Set it to the maximum number of days
      that you anticipate passing without this Agent receiving an incoming Event.
      MD
    end

    event_description <<-MD
      Events look like this:

          {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('XXXXXXXXXXXXXXXXX')/messages/$entity",
            "@odata.etag": "W/\"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\"",
            "id": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "createdDateTime": "2024-09-27T21:29:15Z",
            "lastModifiedDateTime": "2024-09-27T21:29:15Z",
            "changeKey": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "categories": [
          
            ],
            "receivedDateTime": "2024-09-27T21:29:15Z",
            "sentDateTime": "2024-09-27T21:29:09Z",
            "hasAttachments": false,
            "internetMessageId": "<XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX>",
            "subject": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "bodyPreview": "XXXXXXXXXXXXXXX",
            "importance": "normal",
            "parentFolderId": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "conversationId": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "conversationIndex": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "isDeliveryReceiptRequested": null,
            "isReadReceiptRequested": false,
            "isRead": false,
            "isDraft": false,
            "webLink": "https://outlook.live.com/owa/?ItemID=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
            "inferenceClassification": "focused",
            "body": {
              "contentType": "html",
              "content": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX>"
            },
            "sender": {
              "emailAddress": {
                "name": "XXXXXXXXXXXXXXX",
                "address": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
              }
            },
            "from": {
              "emailAddress": {
                "name": "XXXXXXXXXXXXXXX",
                "address": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
              }
            },
            "toRecipients": [
              {
                "emailAddress": {
                  "name": "XXXXXXXXXXXXXXXXXXXXXXXX",
                  "address": "XXXXXXXXXXXXXXXXXXXXXXXX"
                }
              }
            ],
            "ccRecipients": [
              {
                "emailAddress": {
                  "name": "XXXXXXXXXXXXXXXXXXXXXXXX",
                  "address": "XXXXXXXXXXXXXXXXXXXXXXXX"
                }
              }
            ],
            "bccRecipients": [
          
            ],
            "replyTo": [
          
            ],
            "flag": {
              "flagStatus": "notFlagged"
            }
          }
    MD

    def default_options
      {
        'client_id' => '{% credential outlook_client_id %}',
        'client_secret' => '{% credential outlook_client_secret %}',
        'access_token' => '{% credential outlook_access_token %}',
        'refresh_token' => '',
        'folder' => '',
        'debug' => 'false',
        'raw_email' => 'false',
        'emit_events' => 'true',
        'expected_receive_period_in_days' => '2',
      }
    end

    form_configurable :client_id, type: :string
    form_configurable :client_secret, type: :string
    form_configurable :refresh_token, type: :string
    form_configurable :access_token, type: :string
    form_configurable :folder, type: :string
    form_configurable :debug, type: :boolean
    form_configurable :raw_email, type: :boolean
    form_configurable :emit_events, type: :boolean
    form_configurable :expected_receive_period_in_days, type: :string
    form_configurable :type, type: :array, values: ['get_new_emails', 'send_email']
    def validate_options
      errors.add(:base, "type has invalid value: should be 'get_new_emails', 'send_email'") if interpolated['type'].present? && !%w(get_new_emails send_email).include?(interpolated['type'])

      unless options['client_id'].present?
        errors.add(:base, "client_id is a required field")
      end

      unless options['client_secret'].present?
        errors.add(:base, "client_secret is a required field")
      end

      unless options['folder'].present?
        errors.add(:base, "folder is a required field")
      end

      unless options['access_token'].present?
        errors.add(:base, "access_token is a required field")
      end

      unless options['refresh_token'].present?
        errors.add(:base, "refresh_token is a required field")
      end

      if options.has_key?('raw_email') && boolify(options['raw_email']).nil?
        errors.add(:base, "if provided, raw_email must be true or false")
      end

      if options.has_key?('emit_events') && boolify(options['emit_events']).nil?
        errors.add(:base, "if provided, emit_events must be true or false")
      end

      if options.has_key?('debug') && boolify(options['debug']).nil?
        errors.add(:base, "if provided, debug must be true or false")
      end

      unless options['expected_receive_period_in_days'].present? && options['expected_receive_period_in_days'].to_i > 0
        errors.add(:base, "Please provide 'expected_receive_period_in_days' to indicate how many days can pass before this Agent is considered to be not working")
      end
    end

    def working?
      event_created_within?(options['expected_receive_period_in_days']) && !recent_error_logs?
    end

    def receive(incoming_events)
      incoming_events.each do |event|
        interpolate_with(event) do
          log event
          trigger_action
        end
      end
    end

    def check
      trigger_action
    end

    private

    def set_credential(name, value)
      c = user.user_credentials.find_or_initialize_by(credential_name: name)
      c.credential_value = value
      c.save!
    end

    def log_curl_output(code,body)

      log "request status : #{code}"

      if interpolated['debug'] == 'true'
        log "body"
        log body
      end

    end

    def token_refresh()

      uri = URI.parse("https://login.microsoftonline.com/consumers/oauth2/v2.0/token")
      request = Net::HTTP::Post.new(uri)
      request.body = "client_id=#{interpolated['client_id']}&client_secret=#{interpolated['client_secret']}&refresh_token=#{interpolated['refresh_token']}&grant_type=refresh_token&scope=https://graph.microsoft.com/Mail.Read offline_access"
      
      req_options = {
        use_ssl: uri.scheme == "https",
      }
      
      response = Net::HTTP.start(uri.hostname, uri.port, req_options) do |http|
        http.request(request)
      end

      log_curl_output(response.code,response.body)

      payload = JSON.parse(response.body)
      if interpolated['access_token'] != payload['access_token']
        set_credential("outlook_access_token", payload['access_token'])
        if interpolated['debug'] == 'true'
          log "outlook_access_token credential updated"
        end
      end
      current_timestamp = Time.now.to_i
      memory['expires_at'] = payload['expires_in'] + current_timestamp

    end

    def check_token_validity()

      if memory['expires_at'].nil?
        token_refresh()
      else
        timestamp_to_compare = memory['expires_at']
        current_timestamp = Time.now.to_i
#        difference_in_hours = (timestamp_to_compare - current_timestamp) / 3600.0
#        if difference_in_hours < 2
        difference_in_minutes = (timestamp_to_compare - current_timestamp) / 60.0
        if difference_in_minutes < 10
          token_refresh()
        else
          log "refresh not needed"
        end
      end
    end

    def get_email(email_id)
      url = "https://graph.microsoft.com/v1.0/me/messages/#{email_id}"
      if interpolated['raw_email'] == 'true'
        url = url + '/$value'
      end
      uri = URI.parse(url)
      request = Net::HTTP::Get.new(uri)
      request.content_type = "application/json"
      request["Authorization"] = "Bearer  #{interpolated['access_token']}"
      
      req_options = {
        use_ssl: uri.scheme == "https",
      }
      
      response = Net::HTTP.start(uri.hostname, uri.port, req_options) do |http|
        http.request(request)
      end

      log_curl_output(response.code,response.body)

      if interpolated['raw_email'] == 'true'
        result = {}
        result['id'] = email_id
        result['raw_mail'] = Base64.encode64(response.body)
        log "raw_mail"
        log result['raw_mail']

        return result
      else
        payload = JSON.parse(response.body)
        return payload
      end
    end

    def send_email()

      check_token_validity()
      uri = URI.parse("https://graph.microsoft.com/v1.0/me/messages")
      request = Net::HTTP::Post.new(uri)
      request.content_type = "application/json"
      request.body = JSON.dump({
        "subject" => "Did you see last night's game?",
        "importance" => "Low",
        "body" => {
          "contentType" => "HTML",
          "content" => "They were <b>awesome</b>!"
        },
        "toRecipients" => [
          {
            "emailAddress" => {
              "address" => "AdeleV@contoso.com"
            }
          }
        ]
      })
      
      req_options = {
        use_ssl: uri.scheme == "https",
      }
      
      response = Net::HTTP.start(uri.hostname, uri.port, req_options) do |http|
        http.request(request)
      end

      log_curl_output(response.code,response.body)

      payload = JSON.parse(response.body)

      if interpolated['emit_events'] == 'true'
        create_event payload: payload
      end
    end

    def get_new_emails()

      check_token_validity()
      uri = URI.parse("https://graph.microsoft.com/v1.0/me/mailFolders/#{interpolated['folder']}/messages?$format=json&$top=20&$select=id")
      request = Net::HTTP::Get.new(uri)
      request.content_type = "application/json"
      request["Authorization"] = "Bearer  #{interpolated['access_token']}"
      
      req_options = {
        use_ssl: uri.scheme == "https",
      }
      
      response = Net::HTTP.start(uri.hostname, uri.port, req_options) do |http|
        http.request(request)
      end

      log_curl_output(response.code,response.body)

      payload = JSON.parse(response.body)
      if payload != memory['last_status']
        payload['value'].each do |email_data|
          found = false
          if !memory['last_status'].nil? and memory['last_status'].present?
            last_status = memory['last_status']
            if interpolated['debug'] == 'true'
              log "email_data"
              log email_data
            end
            last_status['value'].each do |email_databis|
              if email_data['id'] == email_databis['id']
                found = true
              end
              if interpolated['debug'] == 'true'
                log "found is #{found}!"
                log "email_databis"
                log email_databis
              end
            end
          end
          if found == false
            create_event payload: get_email(email_data['id'])
          else
            if interpolated['debug'] == 'true'
              log "found is #{found}"
            end
          end
        end
        memory['last_status'] = payload
      end

    end

    def trigger_action

      case interpolated['type']
      when "get_new_emails"
        get_new_emails()
      when "send_email"
        send_email()
      else
        log "Error: type has an invalid value (#{interpolated['type']})"
      end
    end
  end
end
