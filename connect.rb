require 'bundler'
Bundler.require


def connect(file_name = '')
  session = GoogleDrive::Session.from_service_account_key('client_secret.json')
  session.spreadsheet_by_key('1ujIZJtVUSXs5RDJ_ZzUPPqc5BbOE4ZMyoG1URg_g38U')
end
